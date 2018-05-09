## TEEST (Text in Excel Every Single Time)

*Disclaimer: Do not use or modify neither the program or the source code to make software violating the law.*

### How do I use it?

1. Head to https://github.com/cozyplanes/teest and download the latest release `EXE` file.
1. Windows may warn you with the missing signature. The file is a DEBUG file, so there isn't a publisher signature. You can proceed downloading anyway since it has been virus checked by the developer.
2. Type the message you want to display in the textbox.
3. Click `Save text` button.
5. To check the file, click `Cancel` button in the opened popup dialog.



### What happens?

When an MS Excel file (`.xlsx`) has been opened, by using TEEST, two files gets opened.

1. The original file user opened
2. Excel file named `message.txt` with the custom message you have written.

`message.txt` excel file will open every single time a person opens a excel file. 

*In some older versions of Excel, the message may overlap with the user opened file.*

### Why does this happen?

When MS Excel program is executed, it is programmed to check the files in the following 2 folders.

- `C:\Program Files\Microsoft Office\Office[versionnumber]\XLSTART`
- `C:\Users\%username%\AppData\Roaming\Microsoft\Excel\XLSTART`

In normal conditions, there is no file in those folders (or the folders doesn't exist at all) but when you use TEEST and click `Save text` button, it saves `message.txt` file in the folders above. From MS Excel is executed again, it will find out there is a file in the folders above, so it will show those text files in Excel.

### Where is this technique used?

There should be a lot of software using this trick, but it is widely known for ransomwares such as `GandCrab` and `TeslaCrypt` displaying decryption methods in MS Excel by this trick.

### How can I disable it?

1. Open TEEST again.
2. Click `Save text` button and click `Cancel` in the following popup.
3. Delete `message.txt` file in the opened explorer.

### LICENSE

This software is under the MIT License. Refer to the `LICENSE` file for more information.

### Contact

<cozyplanes@tuta.io>

Spam/Ads not allowed. Please only send questions or concerns about the software. It may take up to 48 hours to get a reply.
