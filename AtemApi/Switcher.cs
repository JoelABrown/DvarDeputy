using Mooseware.Tachnit.AtemApi;
using System.Collections.Generic;

namespace Mooseware.ClipRunner;

/// <summary>
/// ATEM Switch functionality for the ClipRunner application
/// Wraps functionality in Mooseware.Tachnit.AtemApi.dll
/// </summary>
public class Switcher
{
    /// <summary>
    /// IpAddress of the (connected) ATEM switcher
    /// </summary>
    public string IpAddress { get => _atem.IpAddress; }
    /// <summary>
    /// Whether or not a connection to the ATEM switcher was successfully established
    /// </summary>
    public bool IsReady { get => _atem.IsReady; }
    /// <summary>
    /// The description of the problem (when applicable) attempting to connect to the ATEM switcher
    /// </summary>
    public string NotReadyReason { get => _atem.NotReadyReason; }
    /// <summary>
    /// Name of the ATEM input for camera 1
    /// </summary>
    public string Input1Name { get; private set; }
    /// <summary>
    /// Whether or not the ATEM input for camera 1 exists on the connected ATEM switcher
    /// </summary>
    public bool Input1Ready { get; private set; }
    /// <summary>
    /// Name of the ATEM input for camera 2
    /// </summary>
    public string Input2Name { get; private set; }
    /// <summary>
    /// Whether or not the ATEM input for camera 2 exists on the connected ATEM switcher
    /// </summary>
    public bool Input2Ready { get; private set; }
    /// <summary>
    /// Name of the ATEM input for Media Player 1
    /// </summary>
    public string InputMP1Name { get; private set; }
    /// <summary>
    /// Whether or not the ATEM input for Media Player 1 exists on the connected ATEM switcher
    /// </summary>
    public bool InputMP1Ready { get; private set; }
    /// <summary>
    /// Name of the ATEM input for Media Player 2
    /// </summary>
    public string InputMP2Name { get; private set; }
    /// <summary>
    /// Whether or not the ATEM input for Media Player 2 exists on the connected ATEM switcher
    /// </summary>
    public bool InputMP2Ready { get; private set; }
    /// <summary>
    /// Name of the ATEM input where clips will be shown
    /// </summary>
    public string InputClipName { get; private set; }
    /// <summary>
    /// Whether or not the ATEM input on which clips are meant to be shown exists on the connected ATEM switcher
    /// </summary>
    public bool InputClipReady { get; private set; }

    /// <summary>
    /// Internal reference to the ATEM switcher control interface
    /// </summary>
    private readonly AtemSwitcher _atem;

    /// <summary>
    /// ID of the Camera 1 input
    /// </summary>
    private readonly long _input1Id;
    /// <summary>
    /// ID of the Camera 2 input
    /// </summary>
    private readonly long _input2Id;
    /// <summary>
    /// ID of the MP1 input
    /// </summary>
    private readonly long _inputMp1Id;
    /// <summary>
    /// ID of the MP2 input
    /// </summary>
    private readonly long _inputMp2Id;
    /// <summary>
    /// ID of the input where clips will be shown
    /// </summary>
    private readonly long _inputClipId;

    public Switcher(string atemIpAddress = "192.168.1.240", string input1 = "CAM1", string input2 = "CAM2",
            string inputMp1 = "MP1", string inputMp2 = "MP2", string inputClip = "PC4")
    {
        // Note the input preferences...
        Input1Name = input1;
        Input1Ready = false;    // Until proven otherwise.
        Input2Name = input2;
        Input1Ready = false;
        InputMP1Name = inputMp1;
        InputMP1Ready = false;
        InputMP2Name = inputMp2;
        InputMP2Ready = false;
        InputClipName = inputClip;
        InputClipReady = false;

        // Connect to the _switcher, if it's where it should be...
        _atem = new(atemIpAddress);

        if (_atem.IsReady)
        {
            // Do we have the first input?
            if (_atem.HasInput(input1))
            {
                var input = _atem.GetInputByName(input1);
                if (input is not null)
                {
                    Input1Name = input.ShortName;   // Note using the ATEM version just in case there's a variation in casing.
                    _input1Id = input.InputID;
                    if (input.PortType == AtemSwitcherPortType.External)
                    {
                        Input1Ready = true;
                    }
                }
            }
            // Do we have the second input?
            if (_atem.HasInput(input2))
            {
                var input = _atem.GetInputByName(input2);
                if (input is not null)
                {
                    Input2Name = input.ShortName;   // Note using the ATEM version just in case there's a variation in casing.
                    _input2Id = input.InputID;
                    if (input.PortType == AtemSwitcherPortType.External)
                    {
                        Input1Ready = true;
                    }
                }
            }
            // Is this the MP1 input?
            if (_atem.HasInput(inputMp1))
            {
                var input = _atem.GetInputByName(inputMp1);
                if (input is not null)
                {
                    InputMP1Name = input.ShortName;   // Note using the ATEM version just in case there's a variation in casing.
                    _inputMp1Id = input.InputID;
                    if (input.PortType == AtemSwitcherPortType.MediaPlayerFill)
                    {
                        InputMP1Ready = true;
                    }
                }
            }
            // Is this the MP2 input?
            if (_atem.HasInput(inputMp2))
            {
                var input = _atem.GetInputByName(inputMp2);
                if (input is not null)
                {
                    InputMP2Name = input.ShortName;   // Note using the ATEM version just in case there's a variation in casing.
                    _inputMp2Id = input.InputID;
                    if (input.PortType == AtemSwitcherPortType.MediaPlayerFill)
                    {
                        InputMP2Ready = true;
                    }
                }
            }
            // Is this the Clip input?
            if (_atem.HasInput(inputClip))
            {
                var input = _atem.GetInputByName(inputClip);
                if (input is not null)
                {
                    InputClipName = input.ShortName;   // Note using the ATEM version just in case there's a variation in casing.
                    _inputClipId = input.InputID;
                    if (input.PortType == AtemSwitcherPortType.MediaPlayerFill)
                    {
                        InputClipReady = true;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Show Camera 1 in the program feed and preview Camera 2
    /// </summary>
    public void Take1Preview2()
    {
        // Take 1 / Preview 2, as long as the _switcher and both of these inputs are ready.
        if (IsReady && Input1Ready && Input2Ready)
        {
            _atem.SetProgramInput(_input1Id);
            _atem.SetPreviewInput(_input2Id);
        }
    }

    /// <summary>
    /// Show Camera 2 in the program feed and preview Camera 1
    /// </summary>
    public void Take2Preview1()
    {
        // Take 2 / Preview 1, as long as the _switcher and both of these inputs are ready.
        if (IsReady && Input1Ready && Input2Ready)
        {
            _atem.SetProgramInput(_input2Id);
            _atem.SetPreviewInput(_input1Id);
        }
    }

    /// <summary>
    /// Show Media Player 1 in the program feed and preview Camera 1
    /// </summary>
    public void TakeMP1Preview1()
    {
        // Take MP1 / Preview 1, as long as the _switcher and both of these inputs are ready.
        if (IsReady && InputMP1Ready && Input1Ready)
        {
            _atem.SetProgramInput(_inputMp1Id);
            _atem.SetPreviewInput(_input1Id);
        }
    }

    /// <summary>
    /// Show Media Player 2 in the program feed and preview Camera 1
    /// </summary>
    public void TakeMP2Preview1()
    {
        // Take MP2 / Preview 1, as long as the _switcher and both of these inputs are ready.
        if (IsReady && InputMP2Ready && Input1Ready)
        {
            _atem.SetProgramInput(_inputMp2Id);
            _atem.SetPreviewInput(_input1Id);
        }
    }

    /// <summary>
    /// Show the clip player input by setting it to Preview and then performing a CUT transition
    /// </summary>
    public void TakeClip()
    {
        // Preview tally the Clip and cut it to program, as long as the _switcher and clip input are ready.
        if (IsReady && InputClipReady)
        {
            _atem.SetPreviewInput(_inputClipId);
            _atem.PerformCut();
        }
    }

    /// <summary>
    /// Get a list of short names actually found on the ATEM switcher
    /// </summary>
    public List<string> SwitcherInputShortNames
    {
        get
        {
            return _atem.GetInputShortNames();
        }
    }

    /// <summary>
    /// Get a listing of general properties for the inputs found on the ATEM switcher
    /// </summary>
    public List<string> ListSwitcherInputs
    {
        get
        {
            List<string> results = [];

            foreach (var input in _atem.SwitcherInputs)
            {
                if (input.ShortName != null && input.ShortName.Length > 0)
                {
                    results.Add($"Short name=[{input.ShortName}] InputId=[{input.InputID}] "
                        + $"PortType=[{input.PortType}] ({((long)input.PortType)})");
                }
            }

            return results;
        }
    }

    /// <summary>
    /// Run an ATEM switcher macro
    /// </summary>
    /// <param name="macroName">The name of the macro to be run, as recorded on the ATEM switcher</param>
    public void RunMacro(string macroName)
    {
        if (IsReady)
        {
            _atem.RunMacro(macroName);
        }
    }

    /// <summary>
    /// Retrieves a list of available ATEM macros recorded on the currently connected ATEM switcher
    /// </summary>
    public List<string> MacroNameList
    {
        get
        {
            return _atem.GetMacroNames();
        }
    }

    /// <summary>
    /// Checks to see what is connected to the AUX OUT of the ATEM switcher
    /// </summary>
    /// <returns>Whether the AUX Out is fed by the Program or Preview feed, or something else</returns>
    public string TempGetAuxOut()
    {
        string result = string.Empty;

        if (_atem != null)
        {
            Input auxInput = _atem.GetAuxInput();
            if (_atem.GetAuxInput != null)
            {
                Input pgmOut = _atem.GetInputByName("PGM");
                Input pvwOut = _atem.GetInputByName("PVW");

                if (auxInput.InputID == pgmOut.InputID)
                {
                    result = "AUX=Program";
                }
                else if (auxInput.InputID == pvwOut.InputID)
                {
                    result = "AUX=Preview";
                }
                else
                {
                    result = "AUX=" + auxInput.ShortName;
                }
            }
            else
            {
                result = "NO AUX";
            }
        }

        return result;

    }
}
