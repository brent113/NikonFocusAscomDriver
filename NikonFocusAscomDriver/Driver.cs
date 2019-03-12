//tabs=4
// --------------------------------------------------------------------------------
// ASCOM Focuser driver for Nikon DSLR
//
// Description:	Control Nikon DSLR manual focus motor.
//              Camera must be set to autofocus.
//
// Implements:	ASCOM Focuser interface version: 3
// Author:		BRS
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// 28-Feb-2019	BRS	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//
#define Focuser

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using NikonFocusControl;
using System.Threading;
using System.Threading.Tasks;

namespace ASCOM.Nikon
{
    //
    // Your driver's DeviceID is ASCOM.Nikon.Focuser
    //
    // The Guid attribute sets the CLSID for ASCOM.Nikon.Focuser
    // The ClassInterface/None addribute prevents an empty interface called
    // _Nikon from being created and used as the [default] interface
    //

    /// <summary>
    /// ASCOM Focuser Driver for Nikon DSLR.
    /// </summary>
    [Guid("88d75c85-0493-452f-bfc5-7944693cff11")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Focuser : IFocuserV3
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.Nikon.Focuser";
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM Focuser Driver for Nikon DSLR";

        //internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        //internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        //internal static string comPort; // Variables to hold the currrent device configuration

        // AsyncReset to block the connect function until the camera connects or disconnects
        AutoResetEvent wait;

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal static TraceLogger tl;

        /// <summary>
        /// Private variable to emulate absolute position
        /// </summary>
        private int positionEmulation = -1;
        private int PositionEmulation
        {
            get
            {
                if (positionEmulation == -1)
                    positionEmulation = focusControl.FocusStepMax; // Save for later

                return positionEmulation;
            }
            set
            {
                positionEmulation = value;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Nikon"/> class.
        /// Must be public for COM registration.
        /// </summary>

        private FocusControl focusControl;

        public Focuser()
        {
            tl = new TraceLogger("", "Nikon");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Focuser", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object

            wait = new System.Threading.AutoResetEvent(false);

            try
            {
                focusControl = new FocusControl();
            }
            catch (Exception e)
            {
                tl.LogMessage("Focuser", e.Message);
                throw new ASCOM.InvalidOperationException("Focuser", e);
            }

            //focusControl.DeviceConnected += FocusControl_DeviceConnectedHandler;
            //focusControl.DeviceDisconnected += FocusControl_DeviceDisconnectedHandler;

            tl.LogMessage("Focuser", "Completed initialisation");
        }

        //
        // PUBLIC COM INTERFACE IFocuserV3 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            //if (IsConnected)
            //    System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        // No command strings implemented
        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");

            throw new ASCOM.MethodNotImplementedException("CommandBlind");
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");

            throw new ASCOM.MethodNotImplementedException("CommandBool");
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");

            // No Nikon text commands implemented
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;

            // Disconnect and dispose
            focusControl.Dispose();
            focusControl = null;
        }

        public bool Connected
        {
            get
            {
                LogMessage("Connected Get", "{0}", IsConnected);
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected Set", "{0}", value);
                if (value == IsConnected)
                    return;

                if (value)
                {
                    //connectedState = true; // Set in Nikon driver event handler, with traditional connect, then execute
                    //LogMessage("Connected Set", "Connecting to Nikon camera");

                    connectedState = true; // Move will dynamically connect
                    LogMessage("Connected Set", "Simulating connection to Nikon camera");

                    try
                    {
                        //Task.Run(() => focusControl.Connect());
                    }
                    catch (Exception e)
                    {
                        LogMessage("Connected Set", e.Message);
                    }

                }
                else
                {
                    //connectedState = false; // Set in Nikon driver event handler, with traditional connect, then execute
                    LogMessage("Connected Set", "Disconnecting from Nikon camera");

                    connectedState = false; // Move will dynamically connect
                    LogMessage("Connected Set", "Simulating disconnection from Nikon camera");

                    try
                    {
                        Task.Run(() => focusControl.Disconnect());
                    }
                    catch (Exception e)
                    {
                        LogMessage("Connected Set", e.Message);
                    }

                }
                //wait.WaitOne(6000);
                wait.Reset();
            }
        }

        private void FocusControl_DeviceConnectedHandler(object sender, EventArgs e)
        {
            LogMessage("NikonDriver", "Connected to Nikon camera");
            connectedState = true;
            wait.Set();
            PositionEmulation = focusControl.FocusStepMax;
        }

        private void FocusControl_DeviceDisconnectedHandler(object sender, EventArgs e)
        {
            LogMessage("NikonDriver", "Disconnected from Nikon camera");
            connectedState = false;
            wait.Set();
        }

        public string Description
        {
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

                string driverInfo = "Nikon DSLR focus control. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "3");
                return Convert.ToInt16("3");
            }
        }

        public string Name
        {
            get
            {
                string name = "Nikon Focuser";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IFocuser Implementation

        public bool Absolute
        {
            get
            {
                tl.LogMessage("Absolute Get", true.ToString());
                return true; // This is an absolute focuser emulator
            }
        }

        public void Halt()
        {
            tl.LogMessage("Halt", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Halt");
        }

        public bool IsMoving
        {
            get
            {
                bool isMoving = false;
                try
                {
                    isMoving = focusControl.IsMoving;
                }
                catch (Exception e)
                {
                    tl.LogMessage("Move", e.Message);
                }

                tl.LogMessage("IsMoving Get", isMoving.ToString());
                return isMoving;
            }
        }

        public bool Link
        {
            get
            {
                tl.LogMessage("Link Get", this.Connected.ToString());
                return this.Connected; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
            set
            {
                tl.LogMessage("Link Set", value.ToString());
                this.Connected = value; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
        }

        public int MaxIncrement
        {
            get
            {
                CheckConnected("MaxIncrement");

                //int maxIncrement = focusControl.FocusStepMax;
                int maxIncrement = 10000; // hard coded

                tl.LogMessage("MaxIncrement Get", maxIncrement.ToString());
                return maxIncrement; // Maximum change in one move
            }
        }

        public int MaxStep
        {
            get
            {
                CheckConnected("MaxIncrement");

                int maxStep = 0;
                try
                {
                    maxStep = 2 * focusControl.FocusStepMax;
                }
                catch (Exception e)
                {
                    tl.LogMessage("Move", e.Message);
                }

                tl.LogMessage("MaxStep Get", maxStep.ToString());
                return maxStep; // Maximum extent of the focuser
            }
        }

        public void Move(int Position)
        {
            CheckConnected("Move");

            tl.LogMessage("Move", Position.ToString());

            int moveSteps = Position - PositionEmulation;
            PositionEmulation = Position;
            try
            {
                //focusControl.Move(moveSteps); // Set the focuser position
                focusControl.ConnectAndMove(moveSteps); // Dynamically connect, set the focuser position, and disconnect
            }
            catch (Exception e)
            {
                tl.LogMessage("Move", e.Message);
            }
        }

        public int Position
        {
            get
            {
                CheckConnected("Position");

                // return focuser position
                tl.LogMessage("Position", PositionEmulation.ToString());
                return PositionEmulation;
            }
        }

        public double StepSize
        {
            get
            {
                // focuser does not know how many microns a step is, so must throw PropertyNotImplementedException

                tl.LogMessage("StepSize Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("StepSize", false);
            }
        }

        public bool TempComp
        {
            get
            {
                tl.LogMessage("TempComp Get", false.ToString());
                return false;
            }
            set
            {
                tl.LogMessage("TempComp Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TempComp", false);
            }
        }

        public bool TempCompAvailable
        {
            get
            {
                tl.LogMessage("TempCompAvailable Get", false.ToString());
                return false; // Temperature compensation is not available in this driver
            }
        }

        public double Temperature
        {
            get
            {
                tl.LogMessage("Temperature Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Temperature", false);
            }
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Focuser";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // FocusControl event handlers set this
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}
