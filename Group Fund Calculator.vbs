'Group Fund Member Calculator
' -Calculates how much each member contributed.
' -Looks for file "Group Fund.csv"
' 
' Davenport651 2010

Const ForReading = 1
strInputPath = "Group Fund.csv"
Set objFSO = CreateObject("Scripting.FileSystemObject")  'allows script to access file system
Set objInTxt = objFSO.openTextFile(strInputPath,ForReading)  'opens text file in reading mode
Set dicBank = createobject("Scripting.dictionary") 
'dictionary object will hold individual's name and individual's contribution amount (i.e. Smith,20 )

dim Total 'total balance of account

x = objInTxt.Readline 'reads the first line (header line)...
x = null    'and nullifies the hell outta' it!

Do until objInTxt.AtEndOfStream   'continue through the rest of the text file
	arrLine = objInTxt.ReadLine 'read the next line
	arrLine = split(arrLine,",") 'split the line at each comma
	name = arrLine(0) 'First column is investor's name
	'Following lines ensure no text is present in number fields, then obtains deposit and withdraw amounts
		On Error Resume Next  'forces scripting engine to resume on errors
		if arrLine(3) = "" then 
		'if 4th column (withdraw amount) is blank then withdraw = 0
			withdraw = 0
			else withdraw = cdbl(arrLine(3))
			'if column has any other value, then convert to 'double' type variable
			If NOT err.Number = 0 then
				'If any error occurs by converting this variable to number then...
				withdraw = 0
				err.clear
			End If
		end if
		if arrLine(4) = "" then 
		'if 5th column (deposit amount) is blank then deposit = 0
			deposit = 0
			else deposit = cdbl(arrLine(4))
			'if column has any other value, then convert to 'double' type variable
			If NOT err.Number = 0 then
				'if any error occurs by converting this variable to number then...
				deposit = 0
				err.clear
			End If
		End If
		On Error GOTO 0  'forces scripting engine to stop on errors
	transact = deposit-withdraw
	Total = Total + transact 'adds the current transaction to the total value
	'wscript.echo name & "   " & withdraw & "   " & deposit & "   " & transact & "   " &total   'message line to view all data in this loop
	If dicBank.exists(name) then
		'If the investor's name already exists in the dictionary then...
		strEntry = dicBank(name)  '...place individual's transaction data into strEntry...
		transact = strEntry + transact '...update the individual's transation amount (previous amount + current amount)...
		dicBank.remove(name)   '...and remove the old dictionary data
	End If
	'wscript.echo transact  'message line to view transact amount in this loop
	dicBank.add name, transact 'add name and transact amount to dictionary
Loop

CommonSharedTotal = Total - dicBank("Common")
	'spread "Common" amount among investors; used exclusively to average individual's share
dicBank.remove("Common")

Set dicCommonCalced = createobject("Scripting.dictionary")  'create a new dictionary object to hold final values
For each person in dicBank 'spread common money amongst each investor in dicBank dictionary object
	personPercent = dicBank.item(person)/CommonSharedTotal 
	'each person's share is a percent of their contributions to the "common shared total" as found earlier
	'wscript.echo person & " " & formatpercent(personpercent)  'will echo each person's percent investment
	dicCommonCalced.add person, (Total * personPercent)
	'add person's name and their percent of the total as calculated
Next

dicCommonCalced.add "Total",Total  'add's a total line to the dictionary object
For each key in dicCommonCalced
	'read and echo each line of dictionary object as follows:
	If key = "Total" then ' "total" line grammatically requires a different output message 
		wscript.echo key & " value is $" & ccur(dicCommonCalced.item(key))
	else wscript.echo key & "'s share is $" & ccur(dicCommonCalced.item(key)) & "."
	'Echo results of members and their values as:
	'	"[Member] has contributed $[X]."
	End If
Next
