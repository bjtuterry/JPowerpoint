JPowerpoint
===========
	JPowerpoint是09年时编写的一个生成office powerpoint的Java库。
Example:
===========	

	//...
	PowerPoint ppt = PowerPointHelper.create("test.pptx"); //创建ppt文件
	Slide slide = ppt.addSlide(); //添加幻灯片
	TextBox tb = slide.addTextBox(0, 0, 100, 100); //添加文本框
	Text text = tb.addText("test", false); //添加文本
	text.setFontSize(24); //设置文本样式
	ppt.save();
	//...
