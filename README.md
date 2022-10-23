# COM Port Configuration VBA

<img src="/COM_PORT_CONFIG.bmp" alt="COM_PORT_CONFIG" title="COM Port Configuration" width="80%" height="80%">

## Installing and Testing

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

Download the file [`RIBBON.xml`](RIBBON.xml) from this folder in preparation for use.  

Follow the [instructions](RibbonCustomisation.md) to install the `RIBBON.xml` customisation file.

</p>
</details>

<details><summary>RibbonX Editor Screenshot</summary>
<p>

**Successful Ribbon XML customisation and validation using RibbonX editor**  
  
<img src="/RIBBON/RIBBONX_CONFIG.bmp" alt="RibbonX" title="RibbonX Result" width="80%" height="80%">

</p>
</details>




