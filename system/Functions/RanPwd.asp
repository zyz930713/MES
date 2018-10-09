<%
Function gen_key(digits) 

'Create and define array 
dim char_array(50) 
char_array(0) = "0" 
char_array(1) = "1" 
char_array(2) = "2" 
char_array(3) = "3" 
char_array(4) = "4" 
char_array(5) = "5" 
char_array(6) = "6" 
char_array(7) = "7" 
char_array(8) = "8" 
char_array(9) = "9" 
char_array(10) = "a" 
char_array(11) = "b" 
char_array(12) = "c" 
char_array(13) = "d" 
char_array(14) = "e" 
char_array(15) = "f" 
char_array(16) = "g" 
char_array(17) = "h" 
char_array(18) = "i" 
char_array(19) = "j" 
char_array(20) = "k" 
char_array(21) = "l" 
char_array(22) = "m" 
char_array(23) = "n" 
char_array(24) = "o" 
char_array(25) = "p" 
char_array(26) = "q" 
char_array(27) = "r" 
char_array(28) = "s" 
char_array(29) = "t" 
char_array(30) = "u" 
char_array(31) = "v" 
char_array(32) = "w" 
char_array(33) = "x" 
char_array(34) = "y" 
char_array(35) = "z" 

'Initiate randomize method for default seeding 
randomize 

'Loop through and create the output based on the the variable passed to 
'the function for the length of the key. 
do while len(output) < digits 
num = char_array(Int((35 - 0 + 1) * Rnd + 0)) 
output = output + num 
loop 

'Set return 
gen_key = output 
End Function 

Function gen_num(digits) 

'Create and define array 
dim char_array(10) 
char_array(0) = "0" 
char_array(1) = "1" 
char_array(2) = "2" 
char_array(3) = "3" 
char_array(4) = "4" 
char_array(5) = "5" 
char_array(6) = "6" 
char_array(7) = "7" 
char_array(8) = "8" 
char_array(9) = "9" 

'Initiate randomize method for default seeding 
randomize 

'Loop through and create the output based on the the variable passed to 
'the function for the length of the key. 
do while len(output) < digits 
num = char_array(Int((9 - 0 + 1) * Rnd + 0)) 
output = output + num 
loop 

'Set return 
gen_num = output 
End Function 
%>