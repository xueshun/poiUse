package com.demo;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;
import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class ReadWriteAndDownloadDocServlet extends HttpServlet{
	private URL base = this.getClass().getResource("");
	
	public void doGet(HttpServletRequest request, 
			HttpServletResponse response) throws ServletException, IOException{
		doPost(request,response);
	}
	
	public void doPost(HttpServletRequest request, 
			HttpServletResponse response) throws ServletException, IOException{
		try {
			
			String fileDir = new File(base.getFile(), "../../../../doc/").getCanonicalPath();
			
			System.out.println(fileDir+"        1111111111"+base.getFile());
			//获取模板文件
			File demoFile=new File(fileDir+"/J_XCJC.doc");
			
			FileInputStream in = new FileInputStream(demoFile);
			HWPFDocument hdt = new HWPFDocument(in);
			
			//替换读取到的word模板内容的指定字段
			Range range = hdt.getRange();
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
			for (Map.Entry<String,String> entry:map.entrySet()) {
				range.replaceText(entry.getKey(),entry.getValue());
			}
			
			//输出word内容文件流，提供下载
			response.reset();
            response.setContentType("application/x-msdownload");
            response.addHeader("Content-Disposition", "attachment; filename=\"1.doc\"");
			ByteArrayOutputStream ostream = new ByteArrayOutputStream();
			ServletOutputStream servletOS = response.getOutputStream();
			hdt.write(ostream);
			servletOS.write(ostream.toByteArray());
			servletOS.flush();
			servletOS.close();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			
		}
	}

}
