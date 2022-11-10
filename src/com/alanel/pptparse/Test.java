package com.alanel.pptparse;

import java.io.IOException;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try 
		{
			PptPictExtraction.PptSingleSlidePictExtractor("D:\\eclipse_workspace\\POIDemo\\test.ppt", 0);
		} 
		catch (IOException e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
	}

}
