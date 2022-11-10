package com.alanel.pptparse;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Dimension2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.hemf.usermodel.HemfPicture;

public class PptPictExtraction 
{
	private static void TransEMFtoPNG(String pathname)
	{
		File f = new File(pathname);
    	
		try(FileInputStream fis = new FileInputStream(f))
		{
			HemfPicture emf = new HemfPicture(fis);
			Dimension2D dim = emf.getSize();
			
			int width = Units.pointsToPixel(dim.getWidth());
		    // keep aspect ratio for height
		    int height = Units.pointsToPixel(dim.getHeight());
		    double max = Math.max(width, height);
		    if (max > 1500) 
		    {
		    	width *= 1500/max;
		    	height *= 1500/max;
		    }
		    
		    BufferedImage bufImg = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
		    Graphics2D g = bufImg.createGraphics();
		    g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
		    g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
		    g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
		    g.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
		    emf.draw(g, new Rectangle2D.Double(0,0,width,height));
		    
		    g.dispose();
		    ImageIO.write(bufImg, "PNG", new File(pathname.substring(0, pathname.length() - 4) + ".png"));
			
		} catch (IOException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	
	}
	public static void PptSingleSlidePictExtractor(String pathname, int index) throws IOException 
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
		Dimension pgsize = ppt.getPageSize();
		int idx = 1;
		
		List<HSLFSlide> list = ppt.getSlides();
		
		if(index == 0)
			for(HSLFSlide slide : list)
			{
//				BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
//			    Graphics2D graphics = img.createGraphics();
//			    // clear the drawing area
//			    graphics.setPaint(Color.white);
//			    graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
//			    // render
//			    slide.draw(graphics);
//			    // save the output
//				String expdirName = dirName.substring(0, dirName.length() - 9);
//				expdirName += "\\slides";
//				File expdir = new File(expdirName);
//				expdir.mkdirs();
//			    FileOutputStream sldOut = new FileOutputStream(expdirName + "\\slide-" + idx++ + ".png");
//			    javax.imageio.ImageIO.write(img, "png", sldOut);
//			    sldOut.close();
				
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
				        String saveFilePn = dirName + "\\slide" + sliderNumber 
								+ "_pict" + num++ + type.extension;
				        FileOutputStream out = new FileOutputStream(saveFilePn);
				        
				        out.write(data);
				        out.close();
				        
				        if(type.extension.equals(".emf"))
				        {
				        	TransEMFtoPNG(saveFilePn);
				        }
				    }
				}
			}
		
		else if(index > 0)
		{
			HSLFSlide slide = list.get(index - 1);
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
			        String saveFilePn = dirName + "\\slide" + sliderNumber 
							+ "_pict" + num++ + type.extension;
			        FileOutputStream out = new FileOutputStream(saveFilePn);
			        
			        out.write(data);
			        out.close();
			        
			        if(type.extension.equals(".emf"))
			        {
			        	TransEMFtoPNG(saveFilePn);
			        }
			    }
			}
		}
	}
	
	public static void PptxSingleSlidePictExtractor(String pathname, int index) throws IOException 
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
		if(index == 0)
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
				        String saveFilePn = dirName + "\\slide" + sliderNumber 
								+ "_pict" + pictNumber++ + type.extension;
				        FileOutputStream out = new FileOutputStream(saveFilePn);
				        
				        out.write(data);
				        out.close();
					
				        if(type.extension.equals(".emf"))
				        {
				        	TransEMFtoPNG(saveFilePn);
				        }
					}
				}
			}
		
		else if(index > 0)
		{
			XSLFSlide slide = ppt.getSlides().get(index - 1);
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
			        String saveFilePn = dirName + "\\slide" + sliderNumber 
							+ "_pict" + pictNumber++ + type.extension;
			        FileOutputStream out = new FileOutputStream(saveFilePn);
			        
			        out.write(data);
			        out.close();
			        
			        if(type.extension.equals(".emf"))
			        {
			        	TransEMFtoPNG(saveFilePn);
			        }
				}
			}
		}
		
		ppt.close();
	}
}
