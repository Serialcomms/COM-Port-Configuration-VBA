# Excel Ribbon Customisation

## Using the RibbonX Editor

1. Close all Office documents before continuing
2. Open the required Excel document in the RibbonX Editor
3. Click on the Excel document name on the left hand side
4. From the RibbonX menu, select **Insert > Office 2010+ Office UI Part**
5. Confirm that **customUI14.xml** appears under the document name on the left
6. Select and double-click **customUI14.xml** to open an empty area on the right
7. Copy-and-paste contents of file [**Ribbon.xml**](Ribbon.xml) into the empty area
8. Click on **Validate** from the RibbonX editor
9. Confirm the **Custom XML UI is well formed** message box
10. Click **Save** from the RibbonX editor menu
11. **Close** the RibbonX editor
12. Re-open the saved document in Excel as normal
13. Confirm that a new tab **COM Ports** is present in the Excel Ribbon menu 
14. Check that tab controls are responsive, including clicking button above drop-down
15. COM Port needs to be opened briefly to apply settings - DLL Error 5 if open fails
16. DLL Error 87 in message box indicates that port does not support selected values
