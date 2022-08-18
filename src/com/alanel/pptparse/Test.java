package com.alanel.pptparse;

import java.io.IOException;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try 
		{
			PptTextExtraction.PptSingleSlideTextExtractor("D:\\Qt_workspace\\src\\TeachTool\\TeachTool\\doc\\test.ppt");
		} 
		catch (IOException e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
	}

}
