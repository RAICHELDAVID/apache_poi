<cfscript>
    apachePoi = createObject("java", "org.apache.poi.xssf.usermodel.XSSFWorkbook").init();
    sheet = apachePoi.createSheet("Total sections sheet");

    mainHeadingStyle = apachePoi.createCellStyle();
    contentStyle = apachePoi.createCellStyle();
    firstValueStyle = apachePoi.createCellStyle();
    secondValueStyle=apachePoi.createCellStyle();
    thirdValueStyle=apachePoi.createCellStyle();
    mainHeadingFont = apachePoi.createFont();
    mainHeadingFont.setFontName("Arial Narrow");
    mainHeadingFont.setFontHeightInPoints(12);
    mainHeadingFont.setBold(true);
    mainHeadingFont.setUnderline(1); 

    contentFont = apachePoi.createFont();
    contentFont.setFontName("Arial Narrow");
    contentFont.setFontHeightInPoints(12);

    commonFont=apachePoi.createFont();
    commonFont.setFontName("Arial Narrow");
    commonFont.setFontHeightInPoints(12);
    commonFont.setBold(true);
    mainHeadingStyle.setFont(mainHeadingFont);
    mainHeadingStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").RIGHT);
    contentStyle.setFont(contentFont);
    contentStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").RIGHT);
    firstValueStyle.setFont(contentFont);
    firstValueStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").CENTER);
    secondValueStyle.setFont(commonFont);
    secondValueStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").CENTER);
    thirdValueStyle.setFont(commonFont);
    thirdValueStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").CENTER);

    rowIdx = 0;
    colIdx = 0;

    sheet.setColumnWidth(colIdx, 15 * 256);
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("TOTAL SECTIONS 1-2:");
    cell.setCellStyle(mainHeadingStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Sub Total:");
    cell.setCellStyle(contentStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total DFI % (No Spoils):");
    cell.setCellStyle(contentStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total Cost:");
    cell.setCellStyle(contentStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));
    firstValueStyle.setFillForegroundColor(createObject("java", "org.apache.poi.xssf.usermodel.XSSFColor").init(createObject("java", "java.awt.Color").init(219, 219, 219)));
    firstValueStyle.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").SOLID_FOREGROUND);
    rowIdx = 1;
    colIdx = 3;
    row = sheet.getRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("$0.00");
    cell.setCellStyle(firstValueStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));


    secondValueStyle.setFillForegroundColor(createObject("java", "org.apache.poi.xssf.usermodel.XSSFColor").init(createObject("java", "java.awt.Color").init(252, 229, 205)));
    secondValueStyle.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").SOLID_FOREGROUND);
    rowIdx++;
    row = sheet.getRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("$0.00");
    cell.setCellStyle(secondValueStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    thirdValueStyle.setFillForegroundColor(createObject("java", "org.apache.poi.xssf.usermodel.XSSFColor").init(createObject("java", "java.awt.Color").init(207,226,243)));
    thirdValueStyle.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").SOLID_FOREGROUND);
    rowIdx++;
    row = sheet.getRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("$0.00");
    cell.setCellStyle(thirdValueStyle);
    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    baos = createObject("java", "java.io.ByteArrayOutputStream").init();
    apachePoi.write(baos);
    apachePoi.close();

    theFile = "Total_sections.xlsx";
    cfheader(name="Content-Disposition", value="attachment; filename=#theFile#");
    cfcontent(type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", variable="#baos.toByteArray()#");
</cfscript>
