##VBA Reference

[Visual Basic Documentation](https://msdn.microsoft.com/en-us/library/office/ee861528.aspx)</br>
https://msdn.microsoft.com/en-us/library/32s6akha(v=vs.90).aspx

###IIF

语法:

	IIf(Expression, TruePart, FalsePart)

参数:
	
- `Expression`: 必要参数。用来判断真伪的表达式
- `TruePart`: 必要参数。如果 expr 为 True，则返回这部分的值或表达式
- `FalsePart`: 必要参数。如果 expr 为 False，则返回这部分的值或表达式

功能: 根据表达式的值，返回两部分中的其中一个

###MsgBox

语法:

	MsgBox Prompt[,Buttons][,Title][,Helpfile,Context]

参数:

- `Prompt`: 必选。字符串表达式，显示在对话框中的消息。如果 Prompt 的内容超过一行，则可以在每一行之间用回车符 `Chr(13)` ，换行符 `Chr(10)` 或是回车与换行符的组合 `Chr(13)&Chr(10)`，即 `vbCrLf` 将各行分隔开来
- `Buttons`: 可选。数值表达式，是一些数值的总和，指定所显示的按钮的数目及形式、使用的图标样式（及声音），缺省按钮以及消息框的强制性等。如果省略，则其缺省值为 0。
- `Title`: 可选。字符串表达式，在对话框标题栏中显示的内容。如果省略 Title，则将应用程序标题 App.Title 放在标题栏中
- `Helpfile`: 可选。字符串表达式，用来向对话框提供上下文相关帮助的帮助文件。如果提供了 Helpfile，则也必须提供 Context
- `Context`: 可选。数值表达式，由帮助文件的作者指定给适当的帮助主题的帮助上下文编号。如果提供了 Context，则也必须提供 Helpfile

功能: 弹出一个对话框，等待用户单击按钮，并返回一个 Integer 值表示用户单击了哪一个按钮。