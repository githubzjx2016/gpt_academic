<%-- 
    Document   : report
    Created on : 2020-4-7, 8:41:48
    Author     : wzh
--%>

<%@page import="dao.DatabaseAccess"%>
<%@page import="java.text.DecimalFormat"%>
<%@page import="com.itextpdf.text.Font.FontFamily"%>
<%@page import="java.text.DateFormat"%>
<%@page import="java.text.SimpleDateFormat"%>
<%@page import="com.itextpdf.text.BaseColor.*"%>
<%@page import="com.itextpdf.text.pdf.*"%>
<%@page import="java.sql.*"%>
<%@page import="com.itextpdf.text.pdf.PdfWriter"%>
<%@page import="java.io.*,com.itextpdf.text.*" %>
<%@page import="dao.ITextDemo" %>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Report Page</title>
    </head>
    <body>
        <form action="" method="post">
            <table align="center" id="hisData" style="border-collapse:separate; border-spacing:20px;">
                <h3 align="center">IEC Test Report</h3>

                <%
                    Connection con = DatabaseAccess.getConnection();
                    //连接数据库的五大参数
//                    String driverClass = "com.mysql.cj.jdbc.Driver";
//                    String serverIp = "localhost";
//                    String databaseName = "demo";
//                    String userName = "root";
//                    String pwd = "wzh950526";
//                    String jdbcUrl = "jdbc:mysql://" + serverIp + ":3306/" + databaseName + "?useUnicode=true&characterEncoding=utf8&serverTimezone=GMT";
//                    //String sql = "select * from test_history";
//
//                    //读取JDBC
//                    Class.forName(driverClass);
//                    //链接数据库
//                    Connection con = DriverManager.getConnection(jdbcUrl, userName, pwd);

                    //如果为空，代表当前的状态不是查询，而是查询所有的内容
                    PreparedStatement ps0, ps1, ps2, ps3, ps4, ps5, ps6, ps7, ps8, ps9, ps00;
                    String dissn = request.getParameter("dissn");
                    //String lastModified = (String) session.getAttribute("testDate");
                    String lastModified = (String) request.getParameter("date");
                    String model= (String) request.getParameter("model");
                    
                    ps00 = con.prepareStatement("select * from test_history where sn=? and model=? and last_modified=?");
                    ps00.setString(1, dissn);
                    ps00.setString(2, model);
                    ps00.setString(3, lastModified);                    

                    ps0 = con.prepareStatement("select * from display_info where sn=? and ws_sn= (select WorkStation_Information from test_history where sn=? and model=? and last_modified=?)");
                    ps0.setString(1, dissn);
                    ps0.setString(2, dissn);
                    ps0.setString(3, model);
                    ps0.setString(4, lastModified);

                    //可能where的条件改为sn, model, date三个
                    ps1 = con.prepareStatement("select * from iec_visual_test where vtid= (select Visual_Evaluation_Data from test_history where sn=? and model=? and last_modified=?)");
                    ps1.setString(1, dissn);
                    ps1.setString(3, lastModified);
                    ps1.setString(2, model);

                    //ps2 = con.prepareStatement("select * from workstation_info where sn=(select ws_sn from display_info where sn=?)");
                    ps2 = con.prepareStatement("select location from workstation_info where sn=(select WorkStation_Information from test_history where sn=? and model=? and last_modified=?)");
                    ps2.setString(1, dissn);
                    ps2.setString(3, lastModified);
                    ps2.setString(2, model);

                    ps3 = con.prepareStatement("select * from angular_viewing where avid= (select Angular_View from test_history where sn=? and model=? and last_modified=?)");
                    ps3.setString(1, dissn);
                    ps3.setString(3, lastModified);
                    ps3.setString(2, model);
                    
                    ps4 = con.prepareStatement("select * from basic_luminance where blid= (select Basic_Luminance_Data from test_history where sn=? and model=? and last_modified=?)");
                    ps4.setString(1, dissn);
                    ps4.setString(3, lastModified);
                    ps4.setString(2, model);
                    
                    ps5 = con.prepareStatement("select * from luminance_response where lrid= (select Display_Qa_Grayscale_Luminace from test_history where sn=? and model=? and last_modified=?)");
                    ps5.setString(1, dissn);
                    ps5.setString(3, lastModified);
                    ps5.setString(2, model);
                    
                    ps6 = con.prepareStatement("select * from chromaticity where chid= (select Chromaticity_Evaluation from test_history where sn=? and model=? and last_modified=?)");
                    ps6.setString(1, dissn);
                    ps6.setString(3, lastModified);
                    ps6.setString(2, model);
                    
                    ps7 = con.prepareStatement("select * from luminance_uniformity where luid= (select Luminance_Deviation_Data from test_history where sn=? and model=? and last_modified=?)");
                    ps7.setString(1, dissn);
                    ps7.setString(3, lastModified);
                    ps7.setString(2, model);
                    
                    ps8 = con.prepareStatement("select * from grayscale_resolution where gcid= (select Greyscale_Chromaticity_Evaluation from test_history where sn=? and model=? and last_modified=?)");
                    ps8.setString(1, dissn);
                    ps8.setString(3, lastModified);
                    ps8.setString(2, model);
                    
                    ps9 = con.prepareStatement("SELECT * FROM test_setup where tsid= (select Test_Setup from test_history where sn=? and model=? and last_modified=?);");
                    ps9.setString(1, dissn);
                    ps9.setString(3, lastModified);
                    ps9.setString(2, model);

                    //ResultSet是一个指向数据库的变量，本质上是不保存任何数据的，执行查询
                    ResultSet rs00 = ps00.executeQuery();
                    ResultSet rs0 = ps0.executeQuery();
                    ResultSet rs1 = ps1.executeQuery();
                    ResultSet rs2 = ps2.executeQuery();
                    //ResultSet rs3 = ps3.executeQuery();
                    ResultSet rs4 = ps4.executeQuery();
                    ResultSet rs5 = ps5.executeQuery();
                    //ResultSet rs6 = ps6.executeQuery();
                    //ResultSet rs7 = ps7.executeQuery();
                    ResultSet rs8 = ps8.executeQuery();
                    ResultSet rs9 = ps9.executeQuery();
                    //boolean flag = rs.next();   //判断返回指针是否还能继续往下移动
                    //显示序号
                    ServletContext servletContext = getServletContext();
                    while (rs00.next()) {
                        Document pdfDoc = new Document();
                        // 将要生成的 pdf 文件的路径输出流 
                        // D:/reportPDF/reportIEC62563_A2_ 是在服务器电脑上的存储路径, C:/Users/wzh/Desktop/reportDIN157_ 是在本机上的存储路径
                        // 获取Windows的临时文件夹路径
                        String tempFolderPath = System.getProperty("java.io.tmpdir");
                        String tempFileName = ("reportIEC62563_A2_" + dissn + "_" + model + "_" + lastModified).replaceAll("[/\\\\:*?| ]", "_") + ".pdf";
                        String absolutePath = tempFolderPath + File.separator + tempFileName;
                        FileOutputStream pdfFile = new FileOutputStream(absolutePath);
                        //FileOutputStream pdfFile = new FileOutputStream(new File("D:/reportPDF/reportIEC62563_A2_" + dissn + "_" + model + "_" + lastModified.replaceAll("[/\\\\:*?| ]", "_") + ".pdf"));
                        PdfWriter.getInstance(pdfDoc, pdfFile);

                        pdfDoc.open();  // 打开 Document 文档

                        //解决中文不显示问题
                        BaseFont bfChinese = BaseFont.createFont("STSong-Light","UniGB-UCS2-H", BaseFont.EMBEDDED);
                        //BaseFont baseFont = BaseFont.createFont("C:/Windows/Fonts/SIMYOU.TTF", BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED); 宋体simsun.ttc
                        //Font font = new Font(bfChinese, 15, Font.NORMAL);
                        Font fontChina13Bold = new Font(bfChinese, 13, Font.BOLD);
                        Font fontChina18 = new Font(bfChinese, 18);
                        Font fontChina13 = new Font(bfChinese, 13);
                        String rusFontPath = servletContext.getRealPath("/fonts/segoeui.ttf");
                        BaseFont bfRussian = BaseFont.createFont(rusFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font fontRus13 = new Font(bfRussian, 13);  
                        Font fontRus13Bold = new Font(bfRussian, 13, Font.BOLD);
                        //Paragraph titleParagraph1= new Paragraph("中文显示",fontChina18);
                        Paragraph titleParagraph = new Paragraph("表 A.2\n 基于IEC 62563标准的诊断显示器持久性测试报告", fontChina13Bold);
                        titleParagraph.setAlignment(Element.ALIGN_CENTER);// 居中
                        pdfDoc.add(titleParagraph);

                        // 空格
                        Paragraph blank1 = new Paragraph(" ");
                        //pdfDoc.add(blank1);

                        // 编号
//                        Chunk c1 = new Chunk("Test Report Number：");
//                        Chunk c2 = new Chunk("  "+rs1.getString("test_date").replaceAll("[/\\\\:*?|-]", ""));
//                        Paragraph snoParagraph = new Paragraph();
//                        snoParagraph.add(c1);
//                        snoParagraph.add(c2);
//                        snoParagraph.setAlignment(Element.ALIGN_RIGHT);
//                        pdfDoc.add(snoParagraph);

                        // 空格
                        pdfDoc.add(blank1);
                        
                        //add blank row in the table
                        PdfPCell BlankRow = new PdfPCell();
                        BlankRow.setBorderWidth(0);
                        BlankRow.setColspan(20);
                        BlankRow.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        BlankRow.setPhrase(new Paragraph("\n"));
                        BlankRow.setExtraParagraphSpace(5);

                        // 表格处理
                        PdfPTable table = new PdfPTable(20);//表格列数
                        PdfPTable table2 = new PdfPTable(20);
                        table.setWidthPercentage(100);// 表格宽度为100%
                        table2.setWidthPercentage(100);

                        //给当前的PdfPTable同时设置这两个属性可避免自动分页出现空白的问题
                        table.setSplitLate(false);
                        table.setSplitRows(true);
                        table2.setSplitLate(false);
                        table2.setSplitRows(true);

                        //report details start 
                        //1 General Information
//                                while(rs000.next()){
                        PdfPCell cell1 = new PdfPCell();
                        cell1.setBorderWidth(1);// Border宽度为1
                        cell1.setColspan(20);// 跨4列
                        cell1.setPhrase(new Paragraph("基础信息\n", fontChina13Bold));
                        cell1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        cell1.setExtraParagraphSpace(10);
                        table.addCell(cell1);

//                        PdfPCell reportNum1 = new PdfPCell();
//                        reportNum1.setBorderWidth(1);// Border宽度为1
//                        reportNum1.setColspan(1);// 跨4列
//                        reportNum1.setPhrase(new Paragraph("Test Report Number"));
//                        reportNum1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        reportNum1.setBackgroundColor(new BaseColor(235, 235, 235));
//                        reportNum1.setExtraParagraphSpace(10);
//                        table.addCell(reportNum1);
//                        
//                        PdfPCell reportNum2 = new PdfPCell();
//                        reportNum2.setBorderWidth(1);// Border宽度为1
//                        reportNum2.setColspan(3);// 跨4列
//                        reportNum2.setPhrase(new Paragraph("test No.556666"));
//                        reportNum2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        reportNum2.setExtraParagraphSpace(10);
//                        table.addCell(reportNum2);
                        

//                        PdfPCell testDate1 = new PdfPCell();
//                        testDate1.setBorderWidth(1);// Border宽度为1
//                        testDate1.setColspan(5);// 跨4列
//                        testDate1.setPhrase(new Paragraph("Test Date"));
//                        testDate1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        testDate1.setBackgroundColor(new BaseColor(235, 235, 235));
//                        testDate1.setExtraParagraphSpace(10);
//                        table.addCell(testDate1);
                        
                        if(rs1.next()){
                        
                        String organization = "/";
                        String location = "/";
                        String ws_name = "/";
                        String scenario = "/";
                        if(!rs00.getString("test_setup").contentEquals("0"))
                        {
                            if(rs9.next())
                            {
                            organization = rs9.getString("organization");
                            location = rs9.getString("location");
                            ws_name = rs9.getString("ws_name");
                            scenario = rs9.getString("scenario");
                            }
                        }
                        String operatorName = rs00.getString("operator_name").isEmpty()? "/" : rs00.getString("operator_name");
                        String displayIndex = rs00.getString("displayIndex").isEmpty()? "/" : rs00.getString("displayIndex");
                        String disApp ="/";
                        if(rs00.getString("dodyregion").toLowerCase().contains("non")){
                            disApp = "Reviewing";
                        }
                        else{
                            disApp = "Diagnostic";
                        }
                        PdfPCell testDate2 = new PdfPCell();
                        testDate2.setBorderWidth(1);// Border宽度为1
                        testDate2.setColspan(20);// 跨4列
                        testDate2.setPhrase(new Paragraph("测试日期: "+lastModified+"\n"
                                                            +"测试者: "+operatorName+"\n"
                                                            +"地址: "+location+"\n"
                                                            +"显示器: "+displayIndex+", "+model+", "+dissn+"\n"
                                                            +"运用: "+disApp+","+ws_name, fontChina13));
                        testDate2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        testDate2.setExtraParagraphSpace(10);
                        table.addCell(testDate2);

//                        PdfPCell performedBy1 = new PdfPCell();
//                        performedBy1.setBorderWidth(1);
//                        performedBy1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        performedBy1.setColspan(5);// 跨三列
//                        performedBy1.setPhrase(new Paragraph("Test Performed By"));
//                        performedBy1.setExtraParagraphSpace(10);
//                        performedBy1.setBackgroundColor(new BaseColor(235, 235, 235));
//                        table.addCell(performedBy1);
                        
                        //if(rs9.next()){
//                        String organization = "/";
//                        String location = "/";
//                        String ws_name = "/";
//                        String scenario = "/";
//                        if(!rs00.getString("test_setup").contentEquals("0"))
//                        {
//                            if(rs9.next())
//                            {
//                            organization = rs9.getString("organization");
//                            location = rs9.getString("location");
//                            ws_name = rs9.getString("ws_name");
//                            scenario = rs9.getString("scenario");
//                            }
//                        }
//                        
//                        PdfPCell performedBy2 = new PdfPCell();
//                        performedBy2.setBorderWidth(1);
//                        performedBy2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        performedBy2.setColspan(15);// 跨三列
//                        if(rs00.getString("operator_name")!=null){
//                        performedBy2.setPhrase(new Paragraph(rs00.getString("operator_name")));}
//                        else{
//                            performedBy2.setPhrase(new Paragraph("/"));
//                        }
//                        performedBy2.setExtraParagraphSpace(10);
//                        table.addCell(performedBy2);
                        //}

//                        PdfPCell facility1 = new PdfPCell();
//                        facility1.setBorderWidth(1);
//                        facility1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        facility1.setColspan(5);// 跨三列
//                        facility1.setPhrase(new Paragraph("Facility"));
//                        facility1.setBackgroundColor(new BaseColor(235, 235, 235));
//                        facility1.setExtraParagraphSpace(10);
//                        table.addCell(facility1);
//                        
//                        PdfPCell facility2 = new PdfPCell();
//                        facility2.setBorderWidth(1);
//                        facility2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        facility2.setColspan(15);// 跨三列
//                        facility2.setPhrase(new Paragraph("Name of facility"));
//                        facility2.setExtraParagraphSpace(10);
//                        table.addCell(facility2);

//                        PdfPCell location1 = new PdfPCell();
//                        location1.setBorderWidth(1);
//                        location1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        location1.setColspan(5);// 跨三列
//                        location1.setPhrase(new Paragraph("Location: "));
//                        location1.setBackgroundColor(new BaseColor(235,235,235));
//                        location1.setExtraParagraphSpace(10);
//                        table.addCell(location1);
//                        
//                        
//                        PdfPCell location2 = new PdfPCell();
//                        location2.setBorderWidth(1);
//                        location2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        location2.setColspan(15);// 跨三列
//                        location2.setPhrase(new Paragraph(location));
//                        location2.setExtraParagraphSpace(10);
//                        table.addCell(location2);
//                        
//                        PdfPCell application1 = new PdfPCell();
//                        application1.setBorderWidth(1);
//                        application1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        application1.setColspan(5);// 跨三列
//                        application1.setPhrase(new Paragraph("Application: "));
//                        application1.setBackgroundColor(new BaseColor(235,235,235));
//                        application1.setExtraParagraphSpace(10);
//                        table.addCell(application1);
//                        
//                        
//                        PdfPCell application2 = new PdfPCell();
//                        application2.setBorderWidth(1);
//                        application2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        application2.setColspan(15);// 跨三列
//                        application2.setPhrase(new Paragraph(scenario));
//                        application2.setExtraParagraphSpace(10);
//                        table.addCell(application2);
//                        
//                        //2 Display Information
//                        PdfPCell disInfo = new PdfPCell();
//                        disInfo.setBorderWidth(1);
//                        disInfo.setColspan(20);
//                        disInfo.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        disInfo.setPhrase(new Paragraph("2  Display Information\n", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 15)));
//                        disInfo.setExtraParagraphSpace(10);
//                        disInfo.setBackgroundColor(new BaseColor(235,235,235));
//                        table.addCell(disInfo);
//                        
//                        PdfPCell disnum1 = new PdfPCell();
//                        disnum1.setBorderWidth(1);
//                        disnum1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        disnum1.setColspan(5);// 跨三列
//                        disnum1.setPhrase(new Paragraph("Display Index"));
//                        disnum1.setBackgroundColor(new BaseColor(235,235,235));
//                        disnum1.setExtraParagraphSpace(10);
//                        table.addCell(disnum1);
//                        
//                        PdfPCell disnum2 = new PdfPCell();
//                        disnum2.setBorderWidth(1);
//                        disnum2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        disnum2.setColspan(15);// 跨三列
//                        if(rs00.getString("displayIndex")!=null){
//                        disnum2.setPhrase(new Paragraph(rs00.getString("displayIndex")));
//                        }
//                        else
//                        {disnum2.setPhrase(new Paragraph("/"));}
//                        disnum2.setExtraParagraphSpace(10);
//                        table.addCell(disnum2);
//
//                        PdfPCell sn1 = new PdfPCell();
//                        sn1.setBorderWidth(1);
//                        sn1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        sn1.setColspan(5);// 跨三列
//                        sn1.setPhrase(new Paragraph("S/N"));
//                        sn1.setBackgroundColor(new BaseColor(235,235,235));
//                        sn1.setExtraParagraphSpace(10);
//                        table.addCell(sn1);
//                        
//                        PdfPCell sn2 = new PdfPCell();
//                        sn2.setBorderWidth(1);
//                        sn2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                        sn2.setColspan(15);// 跨三列
//                        sn2.setPhrase(new Paragraph(dissn));
//                        sn2.setExtraParagraphSpace(10);
//                        table.addCell(sn2);
//
//                        if (rs0.next()) {
//                            PdfPCell model1 = new PdfPCell();
//                            model1.setBorderWidth(1);
//                            model1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            model1.setColspan(5);// 跨三列
//                            model1.setPhrase(new Paragraph("Model"));
//                            model1.setBackgroundColor(new BaseColor(235,235,235));
//                            model1.setExtraParagraphSpace(10);
//                            table.addCell(model1);
//                            
//                            PdfPCell model2 = new PdfPCell();
//                            model2.setBorderWidth(1);
//                            model2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            model2.setColspan(15);// 跨三列
//                            model2.setPhrase(new Paragraph(model));
//                            model2.setExtraParagraphSpace(10);
//                            table.addCell(model2);
//
//                            PdfPCell disType1 = new PdfPCell();
//                            disType1.setBorderWidth(1);
//                            disType1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            disType1.setColspan(5);// 跨三列
//                            disType1.setPhrase(new Paragraph("Display Type"));
//                            disType1.setBackgroundColor(new BaseColor(235,235,235));
//                            disType1.setExtraParagraphSpace(10);
//                            table.addCell(disType1);
//                            
//                            PdfPCell disType2 = new PdfPCell();
//                            disType2.setBorderWidth(1);
//                            disType2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            disType2.setColspan(15);// 跨三列
//                            disType2.setPhrase(new Paragraph(rs0.getString("display_type_index")+ " / " + rs0.getString("display_color_index")));
//                            disType2.setExtraParagraphSpace(10);
//                            table.addCell(disType2);
//
//                            PdfPCell manu1 = new PdfPCell();
//                            manu1.setBorderWidth(1);
//                            manu1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            manu1.setColspan(5);// 跨三列
//                            manu1.setPhrase(new Paragraph("Manufacturer"));
//                            manu1.setBackgroundColor(new BaseColor(235,235,235));
//                            manu1.setExtraParagraphSpace(10);
//                            table.addCell(manu1);
//                            
//                            PdfPCell manu2 = new PdfPCell();
//                            manu2.setBorderWidth(1);
//                            manu2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
//                            manu2.setColspan(15);// 跨三列
//                            manu2.setPhrase(new Paragraph(rs0.getString("manu_name")));
//                            manu2.setExtraParagraphSpace(10);
//                            table.addCell(manu2);
//                        }
//                                }


                        //Global Result
                        PdfPCell globalResult = new PdfPCell();
                        globalResult.setBorderWidth(1);
                        globalResult.setColspan(18);
                        globalResult.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        globalResult.setPhrase(new Paragraph("整体测试结果\n", fontChina13Bold));
                        globalResult.setBackgroundColor(new BaseColor(235, 235, 235));
                        globalResult.setExtraParagraphSpace(10);
                        table.addCell(globalResult);

                        PdfPCell globalResult1 = new PdfPCell();
                        globalResult1.setBorderWidth(1);
                        globalResult1.setColspan(2);
                        globalResult1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        String flag1 = "";
                        if ((rs00.getString("test_result").startsWith("1"))) {
                            flag1 = "通过";
                        } else {
                            flag1 = "不通过";
                        }
                        globalResult1.setPhrase(new Paragraph(flag1, fontChina13));
                        table.addCell(globalResult1);
                        

                        //3 Visual Test Headers 
                        PdfPCell vt = new PdfPCell();
                        vt.setBorderWidth(1);
                        vt.setColspan(20);
                        vt.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        vt.setPhrase(new Paragraph("视觉评价\n", fontChina13Bold));
                        vt.setBackgroundColor(new BaseColor(235, 235, 235));
                        vt.setExtraParagraphSpace(5);
                        table.addCell(vt);

                        PdfPCell em = new PdfPCell();
                        em.setBorderWidth(1);
                        em.setColspan(6);
                        em.setRowspan(2);
                        em.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        em.setPhrase(new Paragraph("评估方法", fontChina13));
                        em.setBackgroundColor(new BaseColor(235, 235, 235));
                        em.setExtraParagraphSpace(10);
                        table.addCell(em);

                        PdfPCell pattern = new PdfPCell();
                        pattern.setBorderWidth(1);
                        pattern.setColspan(6);
                        pattern.setRowspan(2);
                        pattern.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        pattern.setPhrase(new Paragraph("设备/工具", fontChina13));
                        pattern.setBackgroundColor(new BaseColor(235, 235, 235));
                        pattern.setExtraParagraphSpace(10);
                        table.addCell(pattern);

                        PdfPCell requirement = new PdfPCell();
                        requirement.setBorderWidth(1);
                        requirement.setColspan(6);
                        requirement.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        requirement.setPhrase(new Paragraph("要求", fontChina13));
                        requirement.setBackgroundColor(new BaseColor(235, 235, 235));
                        requirement.setExtraParagraphSpace(10);
                        table.addCell(requirement);

                        PdfPCell resultHead = new PdfPCell();
                        resultHead.setBorderWidth(1);
                        resultHead.setColspan(2);
                        resultHead.setRowspan(2);
                        resultHead.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        resultHead.setPhrase(new Paragraph("结论", fontChina13));
                        resultHead.setExtraParagraphSpace(10);
                        table.addCell(resultHead);
                        
                        PdfPCell res = new PdfPCell();
                        res.setBorderWidth(1);
                        res.setColspan(6);
                        res.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        res.setPhrase(new Paragraph("测试结果", fontChina13));
                        res.setExtraParagraphSpace(10);
                        table.addCell(res);

                        //3 visual test details
                        PdfPCell oiq1 = new PdfPCell();
                        oiq1.setBorderWidth(1);
                        oiq1.setRowspan(2);
                        oiq1.setColspan(6);
                        oiq1.setVerticalAlignment(PdfPCell.ALIGN_LEFT);
                        oiq1.setPhrase(new Paragraph("整体图像质量评估\n–验证整体性能", fontChina13));
                        oiq1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(oiq1);

                        PdfPCell oiq2 = new PdfPCell();
                        oiq2.setBorderWidth(1);
                        oiq2.setRowspan(2);
                        oiq2.setColspan(6);
                        oiq2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        oiq2.setPhrase(new Paragraph("TG18-QC测试图片", fontChina13));
                        oiq2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(oiq2);

                        PdfPCell oiq3 = new PdfPCell();
                        oiq3.setBorderWidth(1);
                        oiq3.setColspan(6);
                        oiq3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        oiq3.setPhrase(new Paragraph("所有外观正常未发现缺陷", fontChina13));
                        oiq3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(oiq3);

                        PdfPCell oiq4 = new PdfPCell();
                        oiq4.setBorderWidth(1);
                        oiq4.setRowspan(2);
                        oiq4.setColspan(2);
                        oiq4.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        String flag2 = "";
                        if ((rs1.getString("overall_quality").startsWith("1"))) {
                            flag2 = "通过";
                        } else {
                            flag2 = "不通过";
                        }
                        oiq4.setPhrase(new Paragraph(flag2, fontChina13));
                        table.addCell(oiq4);
                        
                        PdfPCell oiq5 = new PdfPCell();
                        oiq5.setBorderWidth(1);
                        oiq5.setColspan(6);
                        oiq5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        oiq5.setPhrase(new Paragraph("无", fontChina13));
                        table.addCell(oiq5);


                        PdfPCell lu1 = new PdfPCell();
                        lu1.setBorderWidth(1);
                        lu1.setRowspan(2);
                        lu1.setColspan(6);
                        lu1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lu1.setPhrase(new Paragraph("照度均匀性评价\n–观察不一致性", fontChina13));
                        lu1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(lu1);

                        PdfPCell lu2 = new PdfPCell();
                        lu2.setBorderWidth(1);
                        lu2.setRowspan(2);
                        lu2.setColspan(6);
                        lu2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lu2.setPhrase(new Paragraph("TG18-UN80测试图片", fontChina13));
                        lu2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(lu2);

                        PdfPCell lu3 = new PdfPCell();
                        lu3.setBorderWidth(1);
                        lu3.setColspan(6);
                        lu3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lu3.setPhrase(new Paragraph("未检测到视觉不均匀", fontChina13));
                        lu3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(lu3);

                        PdfPCell lu4 = new PdfPCell();
                        lu4.setBorderWidth(1);
                        lu4.setRowspan(2);
                        lu4.setColspan(2);
                        lu4.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        String flag5 = "";
                        if ((rs1.getString("luminance_uniformity").startsWith("1"))) {
                            flag5 = "通过";
                        } else {
                            flag5 = "不通过";
                        }
                        lu4.setPhrase(new Paragraph(flag5, fontChina13));
                        table.addCell(lu4);
                        
                        PdfPCell lu5 = new PdfPCell();
                        lu5.setBorderWidth(1);
                        lu5.setColspan(6);
                        lu5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lu5.setPhrase(new Paragraph("无", fontChina13));
                        table.addCell(lu5);


                        PdfPCell cimg1 = new PdfPCell();
                        cimg1.setBorderWidth(1);
                        cimg1.setRowspan(2);
                        cimg1.setColspan(6);
                        cimg1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        cimg1.setPhrase(new Paragraph("临床评估", fontChina13));
                        cimg1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(cimg1);

                        PdfPCell cimg2 = new PdfPCell();
                        cimg2.setBorderWidth(1);
                        cimg2.setRowspan(2);
                        cimg2.setColspan(6);
                        cimg2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        cimg2.setPhrase(new Paragraph("临床测试图片 TG18-CH, TG18-CN, TG18-MM1, TG18-MM2", fontChina13));
                        cimg2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(cimg2);

                        PdfPCell cimg3 = new PdfPCell();
                        cimg3.setBorderWidth(1);
                        cimg3.setColspan(6);
                        cimg3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        cimg3.setPhrase(new Paragraph("临床图片呈现是否OK", fontChina13));
                        cimg3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table.addCell(cimg3);

                        PdfPCell cimg4 = new PdfPCell();
                        cimg4.setBorderWidth(1);
                        cimg4.setRowspan(2);
                        cimg4.setColspan(2);
                        cimg4.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        String flag9 = "";
                        if ((rs1.getString("ImageResult").startsWith("1"))) {
                            flag9 = "通过";
                        } else {
                            flag9 = "不通过";
                        }
                        cimg4.setPhrase(new Paragraph(flag9, fontChina13));
                        table.addCell(cimg4);
                        
                        PdfPCell cimg5 = new PdfPCell();
                        cimg5.setBorderWidth(1);
                        cimg5.setColspan(6);
                        cimg5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        cimg5.setPhrase(new Paragraph("是", fontChina13));
                        table.addCell(cimg5);

                        
                        //table.addCell(BlankRow);
                        
                        //4 Quantitative Test 开始
                        PdfPCell qt = new PdfPCell();
                        qt.setBorderWidth(1);
                        qt.setColspan(20);
                        qt.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        qt.setPhrase(new Paragraph("定量评估\n", fontChina13Bold));
                        qt.setBackgroundColor(new BaseColor(235, 235, 235));
                        qt.setExtraParagraphSpace(5);
                        table2.addCell(qt);

                        PdfPCell bl1 = new PdfPCell();
                        bl1.setBorderWidth(1);
                        bl1.setColspan(6);
                        bl1.setRowspan(2);
                        bl1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        bl1.setPhrase(new Paragraph("基本亮度评估", fontChina13));
                        bl1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(bl1);

                        PdfPCell bl2 = new PdfPCell();
                        bl2.setBorderWidth(1);
                        bl2.setColspan(6);
                        bl2.setRowspan(2);
                        bl2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        bl2.setPhrase(new Paragraph("亮度计\n照度计", fontChina13));
                        bl2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(bl2);

                        PdfPCell bl3 = new PdfPCell();
                        bl3.setBorderWidth(1);
                        bl3.setColspan(6);
                        bl3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        bl3.setPhrase(new Paragraph("r′> 250\na < 0.4\n"));
                        bl3.setExtraParagraphSpace(5);
                        bl3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(bl3);

                        
                        if(rs4.next()){
                        PdfPCell bl4 = new PdfPCell();
                        bl4.setBorderWidth(1);
                        bl4.setColspan(2);
                        bl4.setRowspan(2);
                        bl4.setVerticalAlignment(PdfPCell.ALIGN_TOP);
                        String flag10 = "";
                        Double lmin= (Double.parseDouble(rs4.getString("min_luminance")))- (Double.parseDouble(rs4.getString("ambient_luminance")));
                        Double lmax= (Double.parseDouble(rs4.getString("max_luminance")))- (Double.parseDouble(rs4.getString("ambient_luminance")));
                        Double r = Double.parseDouble(rs4.getString("luminance_ratio"));
                        Double sf = Double.parseDouble(rs4.getString("safety_factor_a"));
                        if(r>250 && sf<0.4){
//                        if ((rs4.getString("result").startsWith("1"))) {
                            flag10 = "通过";
                        } else {
                            flag10 = "不通过";
                        }
                        bl4.setPhrase(new Paragraph(flag10, fontChina13));
                        table2.addCell(bl4);

                        PdfPCell bl5 = new PdfPCell();
                        bl5.setBorderWidth(1);
                        bl5.setColspan(6);
                        bl5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        bl5.setPhrase(new Paragraph("用方法C(B.2.3)测试\n"
                                + "Lmax= " + new DecimalFormat("0.000").format(lmax) + " cd/m2²\n"
                                + "Lmin= " + new DecimalFormat("0.000").format(lmin) + " cd/m2²\n"
                                + "E0= " + new DecimalFormat("0.000").format(rs4.getDouble("illuminance_e")) + " lx\n"
                                + "Rd= " + new DecimalFormat("0.000").format(rs4.getDouble("diffuse_coeff")) + "\n"
                                + "Lamb= " + new DecimalFormat("0.000").format(rs4.getDouble("ambient_luminance")) + " cd/m2²\n"
                                + "r'= " + new DecimalFormat("0.000").format(rs4.getDouble("luminance_ratio")) + "\n"
                                + "a= " + new DecimalFormat("0.000").format(rs4.getDouble("safety_factor_a")) + "\n", fontChina13));
                        bl5.setExtraParagraphSpace(5);
                        table2.addCell(bl5);
                        }


                        PdfPCell lre1 = new PdfPCell();
                        lre1.setBorderWidth(1);
                        lre1.setColspan(6);
                        lre1.setRowspan(2);
                        lre1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lre1.setPhrase(new Paragraph("亮度响应评估", fontChina13));
                        lre1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(lre1);

                        PdfPCell lre2 = new PdfPCell();
                        lre2.setBorderWidth(1);
                        lre2.setColspan(6);
                        lre2.setRowspan(2);
                        lre2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lre2.setPhrase(new Paragraph("亮度计\n照度计", fontChina13));
                        lre2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(lre2);

                        PdfPCell lre3 = new PdfPCell();
                        lre3.setBorderWidth(1);
                        lre3.setColspan(6);
                        lre3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lre3.setPhrase(new Paragraph("最大误差 < 15%\n", fontChina13));
                        lre3.setExtraParagraphSpace(5);
                        lre3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(lre3);

                        if(rs5.next()){
                        PdfPCell lre4 = new PdfPCell();
                        lre4.setBorderWidth(1);
                        lre4.setColspan(2);
                        lre4.setRowspan(2);
                        lre4.setVerticalAlignment(PdfPCell.ALIGN_TOP);
                        String flag11 = "";
                        Double lrMaxDev= Double.parseDouble(rs5.getString("max_deviation"));
                        if (lrMaxDev<15) {
                            flag11 = "通过";
                        } else {
                            flag11 = "不通过";
                        }
                        lre4.setPhrase(new Paragraph(flag11, fontChina13));
                        table2.addCell(lre4);

                        PdfPCell lre5 = new PdfPCell();
                        lre5.setBorderWidth(1);
                        lre5.setColspan(6);
                        lre5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        lre5.setPhrase(new Paragraph("用方法C(B.2.3)测量\n"
                                + "L'(LN01)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln01")) + " cd/m2²\n"
                                + "L'(LN02)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln02")) + " cd/m2²\n"
                                + "L'(LN03)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln03")) + " cd/m2²\n"
                                + "L'(LN04)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln04")) + " cd/m2²\n"
                                + "L'(LN05)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln05")) + " cd/m2²\n"
                                + "L'(LN06)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln06")) + " cd/m2²\n"
                                + "L'(LN07)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln07")) + " cd/m2²\n"
                                + "L'(LN08)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln08")) + " cd/m2²\n"
                                + "L'(LN09)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln09")) + " cd/m2²\n"
                                + "L'(LN10)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln10")) + " cd/m2²\n"
                                + "L'(LN11)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln11")) + " cd/m2²\n"
                                + "L'(LN12)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln12")) + " cd/m2²\n"
                                + "L'(LN13)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln13")) + " cd/m2²\n"
                                + "L'(LN14)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln14")) + " cd/m2²\n"
                                + "L'(LN15)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln15")) + " cd/m2²\n"
                                + "L'(LN16)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln16")) + " cd/m2²\n"
                                + "L'(LN17)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln17")) + " cd/m2²\n"
                                + "L'(LN18)= " + new DecimalFormat("0.000").format(rs5.getDouble("ln18")) + " cd/m2²\n"
                                + "最大误差= " + new DecimalFormat("0.00").format(rs5.getDouble("max_deviation")) + " %\n", fontChina13));
                        lre5.setExtraParagraphSpace(5);
                        table2.addCell(lre5);
                        }
                        
                        //table.addCell(BlankRow);

                        if(rs0.next()) {
                        if(rs0.getString("display_color_index").equals("Color")){
                        PdfPCell gc1 = new PdfPCell();
                        gc1.setBorderWidth(1);
                        gc1.setColspan(6);
                        gc1.setRowspan(2);
                        gc1.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        gc1.setPhrase(new Paragraph("灰度色度评估 \n注:该设备已根据GSDF进行校准", fontChina13));
                        gc1.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(gc1);

                        PdfPCell gc2 = new PdfPCell();
                        gc2.setBorderWidth(1);
                        gc2.setColspan(6);
                        gc2.setRowspan(2);
                        gc2.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        gc2.setPhrase(new Paragraph("色度计", fontChina13));
                        gc2.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(gc2);

                        PdfPCell gc3 = new PdfPCell();
                        gc3.setBorderWidth(1);
                        gc3.setColspan(6);
                        gc3.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        gc3.setPhrase(new Paragraph("最大误差 < 0.01", fontChina13));
                        gc3.setExtraParagraphSpace(5);
                        gc3.setBackgroundColor(new BaseColor(235, 235, 235));
                        table2.addCell(gc3);

                        if(rs8.next()){
                        PdfPCell gc4 = new PdfPCell();
                        gc4.setBorderWidth(1);
                        gc4.setColspan(2);
                        gc4.setRowspan(2);
                        gc4.setVerticalAlignment(PdfPCell.ALIGN_TOP);
                        String flag14 = "";
                        Double gcMaxDev= Double.parseDouble(rs8.getString("max_deviation"));
                        //if(rs8.getString("result").startsWith("1")){
                        if (gcMaxDev<0.01) {
                            flag14 = "通过";
                        } else {
                            flag14 = "不通过";
                        }                          
                        gc4.setPhrase(new Paragraph(flag14, fontChina13));
                        table2.addCell(gc4);

                        double[] lnDb = new double[18];
                        double[] uDb = new double[18];
                        double[] vDb = new double[18];
                        String[] lnStr = new String[18];
                        String[] uStr = new String[18];
                        String[] vStr = new String[18];
                        String[] resStr = new String[18];
                        String[] headStr = new String[18];
                        for(int a=0;a<18;a++){
                            if(a<9){
                                headStr[a] = "LN0"+(a+1)+": ";
                            }
                            else{
                                headStr[a] = "LN"+(a+1)+": ";
                            }
                        }
                        int j=-1;
                        for(int i =0; i<18;i++){
                            uDb[i] = rs8.getDouble(i*3+4);
                            uStr[i] = new DecimalFormat("0.0000").format(uDb[i]);
                            vDb[i] = rs8.getDouble(i*3+5);
                            vStr[i] = new DecimalFormat("0.0000").format(vDb[i]);
                            lnDb[i] = rs8.getDouble(i*3+6);
                            lnStr[i] = new DecimalFormat("0.000").format(lnDb[i]);
                            if(lnDb[i]<5){ 
                                j=i;
                                resStr[i] = headStr[i]+"L= "+lnStr[i]+ " cd/m2²\n"+"u'= "+uStr[i]+ ", v'="+vStr[i]+"\n";
                            }
                            else{
                                if(i-j==1){resStr[i] = "\n剩余测量值:\n"+headStr[i]+"u'= "+uStr[i]+ ", v'="+vStr[i]+"\n";}
                                else{resStr[i] = headStr[i]+"u'= "+uStr[i]+ ", v'="+vStr[i]+"\n";}
                            }
                        }
                        PdfPCell gc5 = new PdfPCell();
                        gc5.setBorderWidth(1);
                        gc5.setColspan(6);
                        gc5.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
                        gc5.setPhrase(new Paragraph("丢弃测量:(L<5 cd/m2²)\n"
                                +resStr[0]+resStr[1]+resStr[2]+resStr[3]+resStr[4]+resStr[5]+resStr[6]+resStr[7]+resStr[8]+resStr[9]+resStr[10]
                                +resStr[11]+resStr[12]+resStr[13]+resStr[14]+resStr[15]+resStr[16]+resStr[17]
                                + "max.Deviation= " + new DecimalFormat("0.0000").format(rs8.getDouble("max_deviation")) + " \n", fontChina13));
                        gc5.setExtraParagraphSpace(5);
                        table2.addCell(gc5);
                        }}
                        }

                        pdfDoc.add(table);
                        pdfDoc.newPage();
                        pdfDoc.add(table2);
                        pdfDoc.close();

                        //ServletContext context = getServletConfig().getServletContext();
                        // D:/reportPDF/reportIEC62563_A2_ 是在服务器电脑上的存储路径, C:/Users/wzh/Desktop/reportDIN157_ 是在本机上的存储路径
                        //文件名考虑使用reportNumber + testType
                        File f = new File(absolutePath);
                        //File f = new File("D:/reportPDF/reportIEC62563_A2_" + dissn + "_" + model + "_" + lastModified.replaceAll("[/\\\\:*?| ]", "_") + ".pdf");
                        String filename = f.getName();
                        //String filename = "reportIEC62563_A2_";  //另存为时的文件名
                        int length = 0;
                        ServletOutputStream op = response.getOutputStream();
                        response.setContentLength((int) f.length());
                        response.setContentType("application/pdf");
                        //判断IE10及以下, IE11的user agent, 使用ie浏览器时直接保存到本地,不在网页进行浏览, 文件名才正确
                        if(request.getHeader("User-Agent").toUpperCase().contains("MSIE")||request.getHeader("User-Agent").toLowerCase().contains("rv:11.0")){
                        response.setHeader("Content-Disposition", "attachment; filename=\"" + filename + "\"");}
                        else{response.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");}
                        response.setHeader("Pragma", "No-cache");
                        response.setHeader("Cache-Control", "No-cache");
                        response.setDateHeader("Expires", 0);

                        byte[] bytes = new byte[1024];
                        DataInputStream in = new DataInputStream(new FileInputStream(f));
                        while ((in != null) && ((length = in.read(bytes)) != -1)) {
                            op.write(bytes, 0, length);
                        }
                        op.close();
                        response.flushBuffer();
                        }
                    } %>
            </table>
            <%
                ps1.close();
                con.close();
            %>
        </form>

    </body>
</html>
