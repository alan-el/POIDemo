package com.alanel.pptparse;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PptPictExtraction 
{
	public static void PptSingleSlidePictExtractor(String pathname) throws IOException 
	{
		String dirName = new String(pathname);
		while(dirName.charAt(dirName.length() - 1) != '.')
		{
			dirName = dirName.substring(0, dirName.length() - 1);
		}
		dirName = dirName.substring(0, dirName.length() - 1);
		
		dirName += "\\pictures";
		
		File dir = new File(dirName);
		
		dir.mkdirs();
		
		HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(pathname));
		List<HSLFSlide> list = ppt.getSlides();
		
		for(HSLFSlide slide : list)
		{
			int sliderNumber = slide.getSlideNumber();
			System.out.println("slide number = " + sliderNumber);
			
			int num = 1;
			for (HSLFShape sh : slide.getShapes()) 
			{
			    if (sh instanceof HSLFPictureShape) 
			    {
			        HSLFPictureShape pict = (HSLFPictureShape) sh;
			        HSLFPictureData pictData = pict.getPictureData();
			        byte[] data = pictData.getData();
			        PictureData.PictureType type = pictData.getType();
			        FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
			        											+ "_pict" + num++ + type.extension);
			        
			        out.write(data);
			        out.close();
			    }
			}
		}
	}
	
	public static void PptxSingleSlidePictExtractor(String pathname) throws IOException 
	{
		String dirName = new String(pathname);
		while(dirName.charAt(dirName.length() - 1) != '.')
		{
			dirName = dirName.substring(0, dirName.length() - 1);
		}
		dirName = dirName.substring(0, dirName.length() - 1);
		
		dirName += "\\pictures";
		
		File dir = new File(dirName);
		
		dir.mkdirs();
		
		FileInputStream is = new FileInputStream(pathname);
		XMLSlideShow ppt = new XMLSlideShow(is);
		is.close();
		
		/* Get the document's all pictures
        for (XSLFPictureData data : ppt.getPictureData()) 
        {
            String type = data.getContentType();
            String name = data.getFileName();
            System.out.println("Picture (" + type + "): " + name);

            InputStream pIs = data.getInputStream();
            // make sense of the image data
            pIs.close();
        }*/
		for (XSLFSlide slide : ppt.getSlides()) 
		{
			int sliderNumber = slide.getSlideNumber();
			System.out.println("slide number = " + sliderNumber);
			int pictNumber = 1;
			
			for (XSLFShape shape : slide) 
			{
				if (shape instanceof XSLFPictureShape) 
				{
					System.out.println("	Shape ID = " + shape.getShapeId());
					XSLFPictureShape pShape = (XSLFPictureShape) shape;
					XSLFPictureData pData = pShape.getPictureData();
					byte[] data = pData.getData();
			        PictureData.PictureType type = pData.getType();
			        FileOutputStream out = new FileOutputStream(dirName + "\\slide" + sliderNumber 
			        											+ "_pict" + pictNumber++ + type.extension);
			        out.write(data);
			        out.close();
				}
			}
		}
		ppt.close();
	}
}
