' getplcrecords
' 
'Script to extract PLC historical data from
'CSV files.
'
'
'Author: Luis Johnson
' ---------------------------------------

' Function Declaration
'

Function GetData(stream)
	'Function receives a stream object obtained 
	'from a CSV files with the following columns:
	' "VarName";"TimeString";"VarValue";"Validity";"Time_ms"
	'The function returns a 5 x counter array where
	'counter is the nymber of lines with valid data

	'Variable to count the number of processed rows 
	Dim counter
	'Variables to hold the column data
	Dim VarName, TimeString, VarValue, Validity, Time_ms
	'Declare a dynamic Array
	ReDim data(5, 10000)
	
	'Initialize line counter
	counter = 0

	'Iterate over the stream and extract the data from every column
	Do While stream.AtEndOfStream <> True 
		line = textStream.ReadLine
		rawData = Split(line, ";")
		VarName = Replace(rawData(0) ,"""","")
		TimeString = Replace(rawData(1) ,"""","")
		VarValue = rawData(2)
		Validity = rawData(3)
		Time_ms = rawData(4)

		'If the data is valid store it into the array and
		'increase the counter
		if Validity = "1" Then
			data(0, counter) = VarName
			data(1, counter) = CDate(Replace(TimeString, ".", "/"))
			data(2, counter) = VarValue
			data(3, counter) = Validity
			data(4, counter) = Time_ms
			counter = counter + 1
		End if
	Loop 

	'Resize the array and preserve the data
	ReDim Preserve data(5, counter)

	'Return the data array
	GetData = data	
End Function

Function FilterData(data, LowerDate, UpperDate)
	'Function receives an array with the following items
	' each one :
	' "VarName";"TimeString";"VarValue";"Validity";"Time_ms"
	'The function return the data betwen the limits 
	'defined bay the parameters LowerDate and UpperDate
	'if there is more than one record in the same hour
	'the function will return the first record every hour'

	'Variable declaration'
	Dim counter, filterCounter,  temp, date, pdate, tdiff 
	
	'declare dynamic array with the same size
	'than the input data array
	reDim fdata(5,Ubound(data,2))

	'Initialize record counter
	filterCounter = 0

	'Initialize variable needed to get one
	'record per hour.'
	tdiff = 1

	'Wscript.echo LowerDate
	For counter = 0 to Ubound(data,2)
		'Get date of the current record
		'for comparasion 
		date = data(1, counter)
		
		'Skip time difference operation on the 
		'first record
		If filterCounter > 0 Then
			pdate = fdata(1, filterCounter - 1)
			tdiff = DateDiff("h", pdate, date)
		End If

		'If statement to check the record data is betwen the limite. 
		'Also, check if the record has the same hour checking the difference 
		'betwen the current record date and the previous record date.
		if date > LowerDate and date < UpperDate and tdiff > 0 Then
			fdata(0, filterCounter) = data(0, counter)
			fdata(1, filterCounter) = data(1, counter)
			fdata(2, filterCounter) = data(2, counter)
			fdata(3, filterCounter) = data(3, counter)
			fdata(4, filterCounter) = data(4, counter)
			filterCounter = filterCounter + 1
		End If
	Next
	ReDim Preserve fdata(5, filterCounter)

	FilterData = fdata
End Function



'Variable Declaration
Dim fileSystemObject, fileObject
Dim textStream
Dim file  
Dim fPath 
Dim line
Dim record()
Dim i
Dim lDate, uDate

'Constant for text file handling
Const ForReading = 1
Const TristateUseDefault = -2

'Scrip argument assigment
Set args =  Wscript.Arguments
fPath = args(0)
lDate = CDate(args(1))
uDate = CDate(args(2))

'System and file Object creation
Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
Set fileObject = fileSystemObject.GetFile(args(0))

'Process input file as text stream
Set textStream = fileObject.OpenAsTextStream(ForReading, TristateUseDefault)

records = GetData(textStream)

filteredData = FilterData(records, lDate, uDate)

For i = 0 to Ubound(filteredData, 2)
	Wscript.echo filteredData(0, i) & "  " & filteredData(1, i)
Next

textStream.close






