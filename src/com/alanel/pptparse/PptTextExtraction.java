package com.alanel.pptparse;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PptTextExtraction 
{
	public static int PptSingleSlideTextExtractor(String pathname, int index) throws IOException 
	{
		
		String dirName = new String(pathname);
		while(dirName.charAt(dirName.length() - 1) != '.')
		{
			dirName = dirName.substring(0, dirName.length() - 1);
		}
		dirName = dirName.substring(0, dirName.length() - 1);
		
		dirName += "\\texts";
		
		File dir = new File(dirName);
		
		dir.mkdirs();
		
		HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(pathname));
		List<HSLFSlide> list = ppt.getSlides();
		
		int slidesNum = list.size();
		
		if(index == 0)
		{

			String sn = String.valueOf(slidesNum);
			FileOutputStream osn = new FileOutputStream(dirName + "\\slides_num.txt");
			osn.write(sn.getBytes());
			osn.close();
			
			for(HSLFSlide slide : list)
			{
				int sliderNumber = slide.getSlideNumber();
				System.out.println("slide number = " + sliderNumber);
				
				List<List<HSLFTextParagraph>> tpLists =  slide.getTextParagraphs();
				int textNumber = 1;
				for(List<HSLFTextParagraph> tpList : tpLists )
				{
					System.out.println("	tpList number = " + tpLists.indexOf(tpList));
	
					String s = HSLFTextParagraph.getText(tpList);
					if(!s.isEmpty())
					{
						FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
																	+ "_text" + textNumber++ + ".txt");
						out.write(s.getBytes());
						out.close();
					}
						
				}
			}
		}
		
		else if(index > 0)
		{
			HSLFSlide slide = list.get(index - 1);
			int sliderNumber = slide.getSlideNumber();
			System.out.println("slide number = " + sliderNumber);
			
			List<List<HSLFTextParagraph>> tpLists =  slide.getTextParagraphs();
			int textNumber = 1;
			for(List<HSLFTextParagraph> tpList : tpLists )
			{
				System.out.println("	tpList number = " + tpLists.indexOf(tpList));

				String s = HSLFTextParagraph.getText(tpList);
				if(!s.isEmpty())
				{
					FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
																+ "_text" + textNumber++ + ".txt");
					out.write(s.getBytes());
					out.close();
				}
					
			}
		}
		
		return list.size();
	}
	
	public static int PptxSingleSlideTextExtractor(String pathname, int index) throws IOException
	{
		String dirName = new String(pathname);
		while(dirName.charAt(dirName.length() - 1) != '.')
		{
			dirName = dirName.substring(0, dirName.length() - 1);
		}
		dirName = dirName.substring(0, dirName.length() - 1);
		
		dirName += "\\texts";
		
		File dir = new File(dirName);
		
		dir.mkdirs();
		
		FileInputStream is = new FileInputStream(pathname);
		XMLSlideShow ppt = new XMLSlideShow(is);
		is.close();
		
		List <XSLFSlide> list = ppt.getSlides();
		
		int slidesNum = list.size();
		
		
		if(index == 0)
		{
			String sn = String.valueOf(slidesNum);
			FileOutputStream osn = new FileOutputStream(dirName + "\\slides_num.txt");
			osn.write(sn.getBytes());
			osn.close();
			
			for (XSLFSlide slide : list) 
	        {
	        	int sliderNumber = slide.getSlideNumber();
				System.out.println("slide number = " + sliderNumber);
	        	int textNumber = 1;
	        	
	            for (XSLFShape shape : slide) 
	            {
	                if (shape instanceof XSLFTextShape) 
	                {
	                	System.out.println("	Shape ID = " + shape.getShapeId());
	                    XSLFTextShape txShape = (XSLFTextShape) shape;
	                    String s = txShape.getText();
	                    
	                    if(!s.isEmpty())
	    				{
	    					FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
	    																+ "_text" + textNumber++ + ".txt");
	    					out.write(s.getBytes());
	    					out.close();
	    				}
	                } 
	//                else if (shape instanceof XSLFPictureShape) 
	//                {
	//                    XSLFPictureShape pShape = (XSLFPictureShape) shape;
	//                    XSLFPictureData pData = pShape.getPictureData();
	//                    System.out.println(pData.getFileName());
	//                } 
	//                else 
	//                {
	//                	System.out.println("Process me: " + shape.getClass());
	//                }
	            }
	        }
		}
		
		else if(index > 0)
		{
			XSLFSlide slide = list.get(index - 1);
			int sliderNumber = slide.getSlideNumber();
			System.out.println("slide number = " + sliderNumber);
        	int textNumber = 1;
        	
            for (XSLFShape shape : slide) 
            {
                if (shape instanceof XSLFTextShape) 
                {
                	System.out.println("	Shape ID = " + shape.getShapeId());
                    XSLFTextShape txShape = (XSLFTextShape) shape;
                    String s = txShape.getText();
                    
                    if(!s.isEmpty())
    				{
    					FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
    																+ "_text" + textNumber++ + ".txt");
    					out.write(s.getBytes());
    					out.close();
    				}
                } 
            }
		}
        
        ppt.close();
        return list.size();
        
        /*
		for (PackagePart p : ppt.getAllEmbeddedParts()) 
		{
            String type = p.getContentType();
            // typically file name
            String name = p.getPartName().getName();
            System.out.println("Embedded file (" + type + "): " + name);

            InputStream pIs = p.getInputStream();
            // make sense of the part data
            pIs.close();
        }
		
		// Get the document's embedded files.
        for (XSLFPictureData data : ppt.getPictureData()) {
            String type = data.getContentType();
            String name = data.getFileName();
            System.out.println("Picture (" + type + "): " + name);

            InputStream pIs = data.getInputStream();
            // make sense of the image data
            pIs.close();
        }

        // size of the canvas in points
        Dimension pageSize = ppt.getPageSize();
        System.out.println("Pagesize: " + pageSize);
		*/
	}
}
