data segment

data ends

code segment
assume cs:code ,ds:data
start:	mov ax,data ;initialise data segment
	mov ds,ax

	int 3
code ends
end start

