using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Nikon;
using Timer = System.Timers.Timer;

namespace NikonFocusControl
{
    public class FocusControl : IDisposable
    {
        // NOTE: For Type0003.md3, the drive state flag is located at index 30 - this might
        //       not be the case for other MD3 files. Please double check your SDK documentation.
        const int driveStateIndex = 38;

        NikonDevice _device;
        NikonManager manager;

        Timer connectionTimer;

        public FocusControl()
        {

        }

        public void Connect()
        {
            connectionTimer = new Timer(5000);
            // Hook up the Elapsed event for the timer. 
            connectionTimer.Elapsed += ConnectionTimerTimedOut;
            connectionTimer.AutoReset = false;
            connectionTimer.Enabled = true;

            try
            {
                // Create manager object - make sure you have the correct MD3 file for your Nikon DSLR (see https://sdk.nikonimaging.com/apply/)
                if (manager != null) try { manager.Shutdown(); } catch { }
                String assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(typeof(FocusControl)).Location) + "\\Type0014.md3";
                manager = new NikonManager(assemblyPath);

                // Handle events, empty handlers will not throw an error on removal
                manager.DeviceAdded -= Manager_DeviceAdded;
                manager.DeviceAdded -= Manager_DeviceAddedNoEvent;
                manager.DeviceRemoved -= Manager_DeviceRemoved;
                manager.DeviceAdded += Manager_DeviceAdded;
                manager.DeviceRemoved += Manager_DeviceRemoved;
            }
            catch
            {
                throw;
            }
        }

        AutoResetEvent wait;
        bool waitTimedOut;
        public void ConnectBlocking()
        {
            connectionTimer = new Timer(5000);
            // Hook up the Elapsed event for the timer. 
            /*connectionTimer.Elapsed += ConnectionTimerTimedOut;
            connectionTimer.AutoReset = false;
            connectionTimer.Enabled = true;*/

            wait = new System.Threading.AutoResetEvent(false);

            try
            {
                Task t = Task.Run(() =>
                {

                    try
                    {
                        // Create manager object - make sure you have the correct MD3 file for your Nikon DSLR (see https://sdk.nikonimaging.com/apply/)
                        if (manager != null) try { manager.Shutdown(); } catch { }
                        String env = "\\Nikon SDK\\" + (Environment.Is64BitProcess ? "x64" : "x86");
                        String assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(typeof(FocusControl)).Location) + env + "\\Type0014.md3";
                        manager = new NikonManager(assemblyPath);

                        // Handle events, empty handlers will not throw an error on removal
                        manager.DeviceAdded -= Manager_DeviceAdded;
                        manager.DeviceAdded -= Manager_DeviceAddedNoEvent;
                        manager.DeviceRemoved -= Manager_DeviceRemoved;
                        manager.DeviceAdded += Manager_DeviceAddedNoEvent;
                    }
                    catch
                    {
                        //wait?.Set();
                        throw;
                    }

                    manager.DeviceRemoved += Manager_DeviceRemoved;
                });

                waitTimedOut = true;
                wait?.WaitOne(5000);
                wait?.Reset();

                if (t.Exception != null)
                {
                    throw t.Exception.InnerException;
                }

                if (waitTimedOut) Connected = false;
            }
            catch
            {
                throw;
            }
        }

        private void ConnectionTimerTimedOut(object sender, ElapsedEventArgs e)
        {
            // Connection timed out.
            DeviceTimedOutDisconnection();
        }

        public void Disconnect()
        {
            try
            {
                _device.LiveViewEnabled = false;
            }
            catch { }

            try
            {
                manager.Shutdown();
            }
            catch { }

            _device = null;
            manager = null;
            IsMoving = false;
            Connected = false;

            if (!disposedValue)
                DeviceDisconnected?.Invoke(this, EventArgs.Empty);
        }

        public void DisconnectAttemptReconnect()
        {
            try
            {
                _device.LiveViewEnabled = false;
            }
            catch { }

            _device = null;
            IsMoving = false;
            Connected = false;

            if (!disposedValue)
                DeviceDisconnected?.Invoke(this, EventArgs.Empty);
        }

        NikonRange driveStep = null;
        public NikonRange DriveStep
        {
            get
            {
                if (driveStep == null)
                {
                    bool tempConnected = Connected; // If not connected, connect and then disconnect
                    try
                    {
                        if (!tempConnected)
                            ConnectBlocking();
                    }
                    catch
                    {
                        throw;
                    }

                    driveStep = GetRange(eNkMAIDCapability.kNkMAIDCapability_MFDriveStep);
                    if (driveStep == null) throw new NullReferenceException("Unable to read Drive Step from device");

                    if (!tempConnected)
                        Disconnect();
                }
                return driveStep;
            }
            set
            {
                if (!Connected)
                    throw new DeviceDisconnectedException("Must be connected to device to set DriveStep");

                driveStep = value;
            }
        }

        public int StepSize
        {
            get
            {
                return (int)DriveStep.Value;
            }
            private set
            {
                if (value < FocusStepMin || value > FocusStepMax) throw new ArgumentOutOfRangeException("Step size must be between min and max.");
                DriveStep.Value = (double)value;

                // Must be in LiveView, send right before in Move
                //SetRange(eNkMAIDCapability.kNkMAIDCapability_MFDriveStep, driveStep);
            }
        }

        public int FocusStepMin
        {
            get
            {
                return (int)DriveStep.Min;
            }
        }

        public int FocusStepMax
        {
            get
            {
                return (int)DriveStep.Max;
            }
        }

        public bool IsMoving { get; private set; }

        public void Move(int steps)
        {
            // Negative is closer, Postive is toward infinity

            if (!Connected)
                throw new DeviceDisconnectedException("Must be connected to device to Move");

            if (steps == 0)
                return;

            eNkMAIDMFDrive dir = steps < 0 ? eNkMAIDMFDrive.kNkMAIDMFDrive_InfinityToClosest : eNkMAIDMFDrive.kNkMAIDMFDrive_ClosestToInfinity;

            StepSize = Math.Abs(steps);

            LiveViewEnabled = true;
            SetRange(eNkMAIDCapability.kNkMAIDCapability_MFDriveStep, DriveStep);
            //Task.Run(() => DriveManualFocus(dir)); // causes a not supported error
            DriveManualFocus(dir);

            LiveViewEnabled = false;

            //Thread.Sleep(3000);
        }

        public void ConnectAndMove(int steps)
        {
            try
            {
                if (!Connected)
                    ConnectBlocking();

                Move(steps);
            }
            catch
            {
                throw;
            }
            finally
            {
                Disconnect();
            }
        }

        public bool Connected { get; private set; }

        #region SDK Device Busy Wrappers
        private NikonRange GetRange(eNkMAIDCapability cap)
        {
            NikonRange range = null;

            if (!Connected) return range;

            bool deviceBusy;
            int attempts = 0;
            Stopwatch stopwatch = Stopwatch.StartNew();
            do
            {
                try
                {
                    deviceBusy = false;
                    attempts++;

                    range = _device.GetRange(cap);
                }
                catch (NikonException ex)
                {
                    deviceBusy = ex.ErrorCode == eNkMAIDResult.kNkMAIDResult_DeviceBusy;

                    if (!deviceBusy)
                        throw;

                    if (stopwatch.ElapsedMilliseconds > 5000)
                        throw new DeviceTimedOutException("Device busy, command timed out");

                    // Exponential delay
                    System.Threading.Thread.Sleep(50 * attempts);
                }
            }
            while (deviceBusy);

            return range;
        }

        private void SetRange(eNkMAIDCapability cap, NikonRange value)
        {
            if (!Connected) return;

            bool deviceBusy;
            int attempts = 0;
            Stopwatch stopwatch = Stopwatch.StartNew();
            do
            {
                try
                {
                    deviceBusy = false;
                    attempts++;

                    _device.SetRange(cap, value);
                }
                catch (NikonException e)
                {
                    deviceBusy = e.ErrorCode == eNkMAIDResult.kNkMAIDResult_DeviceBusy;

                    if (!deviceBusy)
                    {
                        Disconnect();
                        throw;
                    }

                    if (stopwatch.ElapsedMilliseconds > 5000)
                    {
                        DeviceTimedOutDisconnection();
                        return;
                    }

                    // Exponential delay
                    System.Threading.Thread.Sleep(50 * attempts);
                }
            }
            while (deviceBusy);
        }

        private bool LiveViewEnabled
        {
            get
            {
                return _device.LiveViewEnabled;
            }
            set
            {
                if (!Connected) return;

                bool deviceBusy;
                int attempts = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();
                do
                {
                    try
                    {
                        deviceBusy = false;
                        attempts++;

                        _device.LiveViewEnabled = value;
                    }
                    catch (NikonException e)
                    {
                        deviceBusy = e.ErrorCode == eNkMAIDResult.kNkMAIDResult_DeviceBusy;

                        if (!deviceBusy)
                        {
                            Disconnect();
                            throw;
                        }

                        if (stopwatch.ElapsedMilliseconds > 5000)
                        {
                            DeviceTimedOutDisconnection();
                            return;
                        }

                        // Exponential delay
                        System.Threading.Thread.Sleep(50 * attempts);
                    }
                }
                while (deviceBusy);
            }

        }

        private void DriveManualFocus(eNkMAIDMFDrive direction)
        {
            NikonLiveViewImage image = null;
            IsMoving = true;

            if (!Connected) return;

            bool deviceBusy;
            int attempts = 0;
            Stopwatch stopwatch = Stopwatch.StartNew();
            do
            {

                try
                {
                    deviceBusy = false;
                    attempts++;

                    // Start driving the manual focus motor
                    _device.SetUnsigned(eNkMAIDCapability.kNkMAIDCapability_MFDrive, (uint)direction);

                    // Alt wait:
                    image = _device.GetLiveViewImage(); // takes approx the time to move -100ms
                    image = _device.GetLiveViewImage();
                    System.Threading.Thread.Sleep(100);
                }
                catch (NikonException e)
                {
                    deviceBusy = e.ErrorCode == eNkMAIDResult.kNkMAIDResult_DeviceBusy;

                    if (!deviceBusy)
                    {
                        Disconnect();
                        throw;
                    }

                    if (stopwatch.ElapsedMilliseconds > 5000)
                    {
                        DeviceTimedOutDisconnection();
                        return;
                    }

                    // This call acts kinda like a block call, sort of. Not completely
                    image = _device.GetLiveViewImage();
                    image = _device.GetLiveViewImage();

                    // Exponential delay
                    System.Threading.Thread.Sleep(50 * attempts);
                }
            }
            while (deviceBusy);

            IsMoving = false;
        }

        private void DeviceTimedOutDisconnection()
        {
            Disconnect();
        }
        #endregion

        private void Manager_DeviceAdded(NikonManager sender, NikonDevice device)
        {
            if (_device == null)
            {
                // Stop conection timmer, successfully connected;
                connectionTimer.Enabled = false;

                // Save device
                _device = device;

                // Signal that we got a device
                Connected = true;

                // Signal wait release
                waitTimedOut = false;
                wait?.Set();

                // Generate event
                DeviceConnected?.Invoke(this, EventArgs.Empty);
            }
        }


        private void Manager_DeviceAddedNoEvent(NikonManager sender, NikonDevice device)
        {
            if (_device == null)
            {
                // Stop conection timmer, successfully connected;
                connectionTimer.Enabled = false;

                // Save device
                _device = device;

                // Signal that we got a device
                Connected = true;

                // Signal wait release
                waitTimedOut = false;
                wait?.Set();

                // Generate event
                //DeviceConnected?.Invoke(this, EventArgs.Empty);
            }
        }

        private void Manager_DeviceRemoved(NikonManager sender, NikonDevice device)
        {
            DisconnectAttemptReconnect();

            // Signal wait release
            waitTimedOut = false;
            wait?.Set();
        }

        public event EventHandler DeviceConnected;
        public event EventHandler DeviceDisconnected;

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                disposedValue = true;

                if (disposing)
                {
                    Disconnect();
                }

                // Free unmanaged resources (unmanaged objects) and override a finalizer below.
                // Set large fields to null.
                _device = null;
                manager = null;
            }
        }

        // Override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        ~FocusControl()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // Uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }
        #endregion
    }

    class DeviceTimedOutException : Exception
    {
        public DeviceTimedOutException()
        {
        }

        public DeviceTimedOutException(string message) : base(message)
        {
        }

        public DeviceTimedOutException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    class DeviceDisconnectedException : Exception
    {
        public DeviceDisconnectedException()
        {
        }

        public DeviceDisconnectedException(string message) : base(message)
        {
        }

        public DeviceDisconnectedException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

}

