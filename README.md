# COM Port Configuration in Excel

<p float="left">
  <img align="top" src="/Images/COM_PORT_CONFIG.bmp" alt="COM_PORT_CONFIG" title="COM Port Tab and Controls" width="50%" height="50%">
  <img align="top"src="/Images/COM_PROPERTIES.bmp" alt="COM_PROPERTIES" title="COM Port Properties" width="25%" height="25%">
</p>



## Installing VBA Module

####  The main VBA module file should be installed first

<details><summary>VBA Module Installation</summary>
<p>

- Download [COM_PORT_ENUM_SETTINGS.bas](COM_PORT_ENUM_SETTINGS.bas) to a known location on your PC  
- Open a new Excel document   
- Enter the VBA Environment (Alt-F11)  
- From VBA Environment, view the Project Explorer (Control-R)  
- From Project Explorer, right-hand click and select Import File  
- Import the file COM_PORT_ENUM_SETTINGS.bas 
- Check that a new module `COM_PORT_ENUM_SETTINGS` is created and visible in the Modules folder
- VBA6 only - delete `PtrSafe` keyword in function definition   
- Close and return to Excel (Alt-Q)  
- IMPORTANT - save document as type Macro-Enabled with a file name of your choice 

  </p>
  </details>

## Ribbon Customisation

<details><summary>Ribbon Customisation</summary>
<p>

#### Adding custom Ribbon tab

The [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.9.0) is recommended for Ribbon customisation.  

Download and install RibbonX following the instructions provided with it.  

Download the file [`RIBBON.xml`](/RIBBON/Ribbon.xml) in preparation for use.  

Follow the [instructions](/RIBBON/RibbonCustomisation.md) to install the `RIBBON.xml` customisation file.

</p>
</details>

<details><summary>RibbonX Editor Screenshot</summary>
<p>

**Successful Ribbon XML customisation and validation using RibbonX editor**  
  
<img src="/Images/RIBBONX_CONFIG.bmp" alt="RibbonX" title="RibbonX Result" width="80%" height="80%">

</p>
</details>

## Using Tab Controls

<details><summary>Select COM Port</summary>
<p>
  
The text **Select Com Port** is clickable. 

- Clicking it will perform another com port scan and update the drop-down box below it. 
- Text will change to **Detect Com Port** if no ports are available. 
- Mouse hovering over it will show a 'Supertip' message with the :-  

1.   number of Com ports available 
2.   last port scan date and time   

  Com ports are not opened and do not need to be free during a port scan.
 
</p>
</details>

<details><summary>COM Port Drop-Down</summary>
<p>
  
- The drop-down is populated when initially opening the Workbook, or by clicking the button above it.  
- Selecting a drop-down item will refresh the Com Port Settings icon on the right.  
  
</p>
</details>

<details><summary>COM Port Settings</summary>
<p>
  
- Clicking the icon will start the Windows Com Port Config dialogue window.
- New settings can be selected in the dialogue window in preparation for change.
- The selected COM port needs to be opened briefly to apply the changes. 
- A confirmation message box appears prior to opening the port and changing settings.

<details><summary>DLL Errors</summary>
<p>

[DLL Errors](/README_DLL_ERRORS.md) may be returned if the port is unavailable or the settings cannot be applied to it.
  
If COM ports have been added, removed or renumbered, click "Select Com Port" to perform a new port scan.  

Note that some port types (e.g. software virtual com ports) may not support any settings changes from default.   

</p>
</details>

</p>
</details>
