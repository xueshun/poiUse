package com.demo;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class ReadAndWriteDoc
{
	private URL base = this.getClass().getResource("");

	
	/**
	 * demoFile 模板文件
	 * newFile 生成文件
	 * map 要填充的数据
	 * */
	public void writeDoc(File demoFile ,File newFile ,Map<String, String> map)
	{
		try
		{	
			FileInputStream in = new FileInputStream(demoFile);
			HWPFDocument hdt = new HWPFDocument(in);
			// Fields fields = hdt.getFields();
			// 读取word文本内容
			Range range = hdt.getRange();
			// System.out.println(range.text());
			
			// 替换文本内容
			for(Map.Entry<String, String> entry : map.entrySet())
			{
				range.replaceText(entry.getKey(), entry.getValue());
			}
			ByteArrayOutputStream ostream = new ByteArrayOutputStream();
			FileOutputStream out = new FileOutputStream(newFile, true);
			hdt.write(ostream);
			// 输出字节流
			out.write(ostream.toByteArray());
			out.close();
			ostream.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

	public static void main(String[] args)
	{
		ReadAndWriteDoc rawDoc = new ReadAndWriteDoc();
		
		try
		{
			String fileDir = new File(rawDoc.base.getFile(), "../../../../doc/").getCanonicalPath();
			System.out.println(fileDir+"        1111111111"+rawDoc.base.getFile());
			//获取模板文件
			File demoFile=new File(fileDir+"/J_XCJC.doc");
			//创建生成的文件
			File newFile=new File(fileDir+"/1.doc");
			Map<String, String> map = new HashMap<String, String>();
			map.put("$QYMC$", "xx数码科技股份有限公司");
			map.put("$QYDZ$", "广东省广州市天河区xx路xx号");
			map.put("$QYFZR$", "张三");
			map.put("$FRDB$", "李四");
			map.put("$CJSJ$", "2000-11-10");
			map.put("$SCPZMSJWT$", "5");
			map.put("$XCJCJBQ$", "6");
			map.put("$JLJJJFF$", "7");
			map.put("$QYFZRQM$", "张三");
			map.put("$CPRWQM$", "赵六");
			map.put("$ZFZH$", "100001");
			map.put("$BZ$", "无");
			rawDoc.writeDoc(demoFile,newFile,map);
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
}
