﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using CameraControl;
using CameraControl.Devices;
using CameraControl.Devices.Nikon;
using CameraControl.Devices.Classes;
using Timer = System.Timers.Timer;

namespace NikonFocusControl
{
    public class FocusControl : IDisposable
    {
        public const int MAX_STEP = 32768;
        public const int MIN_STEP = -32768;

        ICameraDevice CameraDevice;
        CameraDeviceManager DeviceManager = new CameraDeviceManager();
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
                // Handle events, empty handlers will not throw an error on removal
                DeviceManager.CameraConnected -= Manager_DeviceAdded;
                DeviceManager.CameraConnected -= Manager_DeviceAddedNoEvent;
                DeviceManager.CameraDisconnected -= Manager_DeviceRemoved;
                DeviceManager.CameraConnected += Manager_DeviceAdded;
                DeviceManager.CameraDisconnected += Manager_DeviceRemoved;

                DeviceManager.DetectWebcams = false;
                DeviceManager.ConnectToCamera();
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
                        // Handle events, empty handlers will not throw an error on removal
                        DeviceManager.CameraConnected -= Manager_DeviceAdded;
                        DeviceManager.CameraConnected -= Manager_DeviceAddedNoEvent;
                        DeviceManager.CameraDisconnected -= Manager_DeviceRemoved;
                        DeviceManager.CameraConnected += Manager_DeviceAdded;
                        DeviceManager.CameraDisconnected += Manager_DeviceRemoved;

                        DeviceManager.DetectWebcams = false;
                        DeviceManager.ConnectToCamera();
                    }
                    catch
                    {
                        //wait?.Set();
                        throw;
                    }
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
                if (CameraDevice != null) DeviceManager.DisconnectCamera(CameraDevice);
            }
            catch { }

            CameraDevice = null;
            DeviceManager.CloseAll();
            IsMoving = false;
            Connected = false;

            if (!disposedValue)
                DeviceDisconnected?.Invoke(this, EventArgs.Empty);
        }

        public int FocusStepMin
        {
            get
            {
                return MIN_STEP;
            }
        }

        public int FocusStepMax
        {
            get
            {
                return MAX_STEP;
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

            LiveViewData image;

            LiveViewEnabled = true; Thread.Sleep(100);
            image = CameraDevice.GetLiveViewImage(); Thread.Sleep(100);
            CameraDevice.GetProhibitionCondition(OperationEnum.ManualFocus);
            CameraDevice.Focus(steps); Thread.Sleep(100);
            image = CameraDevice.GetLiveViewImage(); Thread.Sleep(100);
            image = CameraDevice.GetLiveViewImage(); Thread.Sleep(100);
            LiveViewEnabled = false; Thread.Sleep(100);
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

        private bool _LiveViewEnabled;
        private bool LiveViewEnabled
        {
            get
            {
                return _LiveViewEnabled;
            }
            set
            {
                if (!Connected) return;

                const int DONE = 10;
                int retryCount = 0;
                do
                {
                    try
                    {
                        if (value)
                        {
                            CameraDevice.StartLiveView();
                        }
                        else
                        {
                            CameraDevice.StopLiveView();
                        }

                        _LiveViewEnabled = value;
                        retryCount = DONE;
                    }
                    catch (DeviceException exception)
                    {
                        if (exception.ErrorCode == ErrorCodes.MTP_Device_Busy || exception.ErrorCode == ErrorCodes.ERROR_BUSY)
                        {
                            Thread.Sleep(100);
                            retryCount++;
                        }
                        else
                        {
                            throw exception;
                        }
                    }
                } while (retryCount < DONE);
            }
        }

        private void DeviceTimedOutDisconnection()
        {
            Disconnect();
        }

        private void Manager_DeviceAdded(ICameraDevice cameraDevice)
        {
            // Ignore virtual camera. Maybe add Canon later/
            if (cameraDevice.Manufacturer != "Nikon Corporation") return;

            if (CameraDevice == null)
            {
                // Stop conection timmer, successfully connected;
                connectionTimer.Enabled = false;

                // Save device
                CameraDevice = cameraDevice;

                // Signal that we got a device
                Connected = true;

                // Signal wait release
                waitTimedOut = false;
                wait?.Set();

                // Generate event
                DeviceConnected?.Invoke(this, EventArgs.Empty);
            }
        }

        private void Manager_DeviceAddedNoEvent(ICameraDevice cameraDevice)
        {
            if (CameraDevice == null)
            {
                // Stop conection timmer, successfully connected;
                connectionTimer.Enabled = false;

                // Save device
                CameraDevice = cameraDevice;

                // Signal that we got a device
                Connected = true;

                // Signal wait release
                waitTimedOut = false;
                wait?.Set();

                // Generate event
                //DeviceConnected?.Invoke(this, EventArgs.Empty);
            }
        }

        private void Manager_DeviceRemoved(ICameraDevice cameraDevice)
        {
            if (CameraDevice != cameraDevice) return;

            IsMoving = false;
            Connected = false;

            if (!disposedValue)
                DeviceDisconnected?.Invoke(this, EventArgs.Empty);

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
                CameraDevice = null;
                DeviceManager = null;
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

