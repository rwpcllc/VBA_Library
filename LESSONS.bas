"'DICTIONARY HASH MAP VALUES VS REFERENCES
'
'Discovered that when creating a HASH table with key value pairs, when adding a key with a cells(#,#) argument as the value, the dictionary adds the key, then adds
'a pointer reference to the cell address instead of the value of the cell contents.   For example, A1 = 10, adding to the dictionary {KEY, cells(1,1)} stores as
'KEY = pointer to A1.  So if for some reason the sheet is deleted later on in the program, A1 becomes an invalid reference.
'SOLUTION: Store the cell value in a variable first and pass in the variable as the value in a key value pair, or DONT delete the sheet Asshole!
'NOTE: If the value in the key value pair is updated (by adding or subtracting to the value) it will no longer store as a reference to the cell.
"
