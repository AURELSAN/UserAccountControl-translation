Simple VBA code to translate UserAccountControl attributes in Excel file from their decimal value to a human readable value.
This code is designed to work best with CSV/Excel template report from Detack appliance.
You can customize the header search by changing the actual column header name you are looking for, by renaming "spare_big_1".

To better understand what UAC Property flags mean please refer to well documented Microsoft site:
https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/useraccountcontrol-manipulate-account-properties
Some other references can be more detailed and applicable to your daily work, like:
http://www.selfadsi.org/ads-attributes/user-userAccountControl.htm

Quick note about flag PASSWD_NOTREQD:
This flag is often misunderstood. If means that the user is not subject to a possibly existing policy regarding the length of password. 
So, the user can have a shorter password than required or even an empty password.

How to use:<br />
Open the desired CSV/Excel file that contains your UAC values in a column with Excel<br />
Select View -> Macros<br />
Type "UAC" in the "Macro name" field (without the "")<br />
Then select "Create"<br />
Copy paste the full code, make sure to erase the "Sub UAC() End Sub" already present<br />
Select depending on your habits "Run"->"Run Sub" / "F5" / the green play button<br />
The results will be written in the column right after the column containing the UAC values (expect 1 second per 200 lines)<br />
You can then save your Excel file, there is no need to register the macro unless you want to do so<br />
