# VBScript for adding Japan new Holidays to Outlook

[日本語版 README はこちら](https://github.com/Microsoft/AddOutlookJapanHoliday/tree/master/ja-jp)

Starting from 2019, there are some additions/changes to Public Holidays in Japan.  

Mainstream support products such as Outlook 2016 / 2019 and Office 365 ProPlus will get this update through Office update. Extended support products such as Outlook 2013 and other legacy versions are excluded from this update, however, we have prepared a script 'AddHolidays.vbs' for Outlook 2010 / 2013 users.

This script will only add new Holidays that are not yet added to the default calendar. If Holidays already exist in the calendar, then the script will not make any changes. This is to prevent creating redundant items.

## How to use

1. Download [AddHolidays.zip](https://github.com/Microsoft/AddOutlookJapanHoliday/releases) and extract it. If you are using Outlook in English and the names of Holidays are English, download AddHolidays_en.zip.
2. Double-click on 'AddHolidays.vbs' (or 'AddHolidays_en.vbs') on a machine running Outlook 2010 / 2013.
3. Open Outlook to confirm Holidays are successfully updated.

## Feedback

If you have any feedback, please post on the [Issues](https://github.com/Microsoft/AddOutlookJapanHoliday/issues) list.

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit <https://cla.microsoft.com>.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.