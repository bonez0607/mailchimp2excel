# Mailchimp2Excel 

## Built With
Visual Basic for Applications (VBA)

## Installation

1. [Download modules](https://github.com/bonez0607/mailchimp2excel/tree/master/Modules)
2. Open Microsoft Excel
3. Open **Visual Basic** editor by pressing `Alt + F11` or selecting it from the [developer tab](https://support.office.com/en-us/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45) 
4. **File > Import File...**
	* Import `JSONConverter.bas` Module
	* Import `Base64Encoder.bas` Module
	* Import `Mailchimp2Excel.bas` Module
##Usage
* Select the sheet you would like to import your subscriber data into from the **VBAProject** window
* Within a `Sub` call the `get_list()` sub. 
 ```
  Sub your_sub_name()
    Call get_list("[YOUR API KEY]", "[YOUR LIST ID]", 100, "[YOUR SHEET NAME]"
  EndSub
 ```
* The parameter that contains `100` is the total number of subscribers you would like to import. 
* Your data should now be displayed in the appropriate sheet!

## Prerequisites
* Microsoft Excel 
* MailChimp Account
* When saving workbook be sure to select **Excel Macro-Enabled Workbook** from `save-as` file type dropdown.
* Mailchimp API key
* List id

## Resources
* [Getting API key](https://kb.mailchimp.com/integrations/api-integrations/about-api-keys)
* [MailChimp API documentation](https://developer.mailchimp.com/documentation/mailchimp/guides/get-started-with-mailchimp-api-3/)
* [Show the developer tab](https://support.office.com/en-us/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)
* [VBA-JSON documentation and installation](https://github.com/VBA-tools/VBA-JSON#installation)
* [base64 encode script](https://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript/506992#506992)

## License
This project is licensed under the [MIT license]("https://opensource.org/licenses/mit-license.php").