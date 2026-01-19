using BMDSwitcherAPI;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Mooseware.Tachnit.AtemApi;

internal class AtemSwitcherDep(IBMDSwitcher switcher)
{
    private readonly IBMDSwitcher switcher = switcher;

    public IEnumerable<IBMDSwitcherMixEffectBlock> MixEffectBlocks
    {
        get
        {
            // Create a mix effect block iterator
            switcher.CreateIterator(typeof(IBMDSwitcherMixEffectBlockIterator).GUID, out IntPtr meIteratorPtr);
            if (Marshal.GetObjectForIUnknown(meIteratorPtr) is not IBMDSwitcherMixEffectBlockIterator meIterator)
            {
                yield break;
            }

            // Iterate through all mix effect blocks
            while (true)
            {
                meIterator.Next(out IBMDSwitcherMixEffectBlock me);
                if (me != null)
                {
                    yield return me;
                }
                else
                {
                    yield break;
                }
            }
        }
    }

    public IEnumerable<IBMDSwitcherInput> SwitcherInputs
    {
        get
        {
            // Create an input iterator
            switcher.CreateIterator(typeof(IBMDSwitcherInputIterator).GUID, out IntPtr inputIteratorPtr);
            if (Marshal.GetObjectForIUnknown(inputIteratorPtr) is not IBMDSwitcherInputIterator inputIterator)
            {
                yield break;
            }
            // Scan through all inputs
            while (true)
            {
                inputIterator.Next(out IBMDSwitcherInput input);
                if (input != null)
                {
                    yield return input;
                }
                else
                {
                    yield break;
                }
            }
        }
    }

    // Can clean this up (it was temp code for poking at the ATEM API for exploratory purposes...
    //public string TempGetAuxOut()
    //{
    //    string result;  // = string.Empty;
    //
    //    IBMDSwitcherInputAux switcherInputAux = (IBMDSwitcherInputAux)_switcher;
    //    switcherInputAux.GetInputSource(out long inputType);
    //    result = inputType.ToString();
    //
    //    IBMDSwitcherInput switcherInput = (IBMDSwitcherInput)_switcher;
    //    switcherInput.GetAvailableExternalPortTypes(out _BMDSwitcherExternalPortType type);
    //    result = type.ToString();
    //
    //    return result;
    //}

    public void RunMacro(uint macroIndex)
    {
        // Use the index
        IBMDSwitcherMacroControl switcherMacroControl = (IBMDSwitcherMacroControl)switcher;
        switcherMacroControl.Run(macroIndex);
    }

    public List<Macro> Macros
    {
        get
        {
            List<Macro> results = [];

            IBMDSwitcherMacroPool switcherMacroPool = (IBMDSwitcherMacroPool)switcher;
            switcherMacroPool.GetMaxCount(out uint macroCount);

            for (uint i = 0; i < macroCount; i++)
            {
                switcherMacroPool.GetName(i, out string macroName);
                if (macroName == null || macroName.Length == 0)
                {
                    macroName = "{null}";
                }
                switcherMacroPool.GetDescription(i, out string macroDescription);
                macroDescription ??= string.Empty;
                if (macroName != "{null}")
                {
                    Macro macro = new()
                    {
                        Index = i,
                        Name = macroName,
                        Description = macroDescription
                    };

                    results.Add(macro);
                }
            }

            return results;
        }
    }
}

