@256
D=A
@SP
M=D
@Sys$Ret.1
D=A
@SP
M=M+1
A=M-1
M=D
@LCL
D=M
@SP
M=M+1
A=M-1
M=D
@ARG
D=M
@SP
M=M+1
A=M-1
M=D
@THIS
D=M
@SP
M=M+1
A=M-1
M=D
@THAT
D=M
@SP
M=M+1
A=M-1
M=D
@5
D=A
@SP
D=M-D
@ARG
M=D
@SP
D=M
@LCL
M=D
@Sys.init
0;JMP
(Sys$Ret.1)
//push argument 1
@ARG
D=M
@1
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
//pop pointer 1
@SP
AM=M-1
D=M
@THAT
M=D
//push constant 0
@0
D=A
@SP
M=M+1
A=M-1
M=D
//pop that 0
@THAT
D=M
@0
D=D+A
@R13
M=D
@SP
AM=M-1
D=M
@R13
A=M
M=D
//push constant 1
@1
D=A
@SP
M=M+1
A=M-1
M=D
//pop that 1
@THAT
D=M
@1
D=D+A
@R13
M=D
@SP
AM=M-1
D=M
@R13
A=M
M=D
//push argument 0
@ARG
D=M
@0
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
//push constant 2
@2
D=A
@SP
M=M+1
A=M-1
M=D
//sub
@SP
AM=M-1
D=M
A=A-1
M=M-D
//pop argument 0
@ARG
D=M
@0
D=D+A
@R13
M=D
@SP
AM=M-1
D=M
@R13
A=M
M=D
(MAIN_LOOP_START)
//push argument 0
@ARG
D=M
@0
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
@SP
AM=M-1
D=M
@COMPUTE_ELEMENT
D;JNE
@END_PROGRAM
0;JMP
(COMPUTE_ELEMENT)
//push that 0
@THAT
D=M
@0
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
//push that 1
@THAT
D=M
@1
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
//add
@SP
AM=M-1
D=M
A=A-1
M=D+M
//pop that 2
@THAT
D=M
@2
D=D+A
@R13
M=D
@SP
AM=M-1
D=M
@R13
A=M
M=D
//push pointer 1
@THAT
D=M
@SP
M=M+1
A=M-1
M=D
//push constant 1
@1
D=A
@SP
M=M+1
A=M-1
M=D
//add
@SP
AM=M-1
D=M
A=A-1
M=D+M
//pop pointer 1
@SP
AM=M-1
D=M
@THAT
M=D
//push argument 0
@ARG
D=M
@0
A=D+A
D=M
@SP
M=M+1
A=M-1
M=D
//push constant 1
@1
D=A
@SP
M=M+1
A=M-1
M=D
//sub
@SP
AM=M-1
D=M
A=A-1
M=M-D
//pop argument 0
@ARG
D=M
@0
D=D+A
@R13
M=D
@SP
AM=M-1
D=M
@R13
A=M
M=D
@MAIN_LOOP_START
0;JMP
(END_PROGRAM)
