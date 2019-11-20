package convertcsv;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class csvToExcel {

    public static void main(String args[]){

        int FileCount=args.length;
        String resultFilename=args[FileCount-1];
        BufferedReader br = null;

        //Create The Excel WorkBook
        Workbook workbook = new XSSFWorkbook();


        //set The cell Style for header
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        headerFont.setFontHeightInPoints((short) 11);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        //set The cell Style for header
        Font normalFont = workbook.createFont();
        normalFont.setBold(false);
        normalFont.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        normalFont.setFontHeightInPoints((short) 11);
        normalFont.setColor(IndexedColors.BLACK.getIndex());

        //set The cell Style for header cell
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        //set The cell Style for normal cells
        CellStyle normalStyle = workbook.createCellStyle();
        normalStyle.setFont(normalFont);
        normalStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        normalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        normalStyle.setBorderBottom(BorderStyle.THIN);
        normalStyle.setBorderTop(BorderStyle.THIN);
        normalStyle.setBorderLeft(BorderStyle.THIN);
        normalStyle.setBorderRight(BorderStyle.THIN);

        // creates the ExcelFile
        FileOutputStream fileOut = null;
        try {

            if(resultFilename.endsWith(".xlsx")){
                fileOut = new FileOutputStream(resultFilename);

            }
            else{
                fileOut = new FileOutputStream(resultFilename+".xlsx");

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }


        //Iterates through every files
        for(int createSheets=0;createSheets<(FileCount-1);createSheets++) {

            String csvFile = args[createSheets];
            Sheet sheet = workbook.createSheet(args[createSheets].split("\\.")[0]);
            String line = "";
            String cvsSplitBy = ",";
            int headerflag = 0;
            try {


                int lineCount = 0;
                int columnCount=0;
                int TotalColumnCount=0;

                br = new BufferedReader(new FileReader(csvFile));
                while ((line = br.readLine()) != null) {
                    columnCount=line.split(cvsSplitBy).length;
                    if (headerflag == 1) {

                        Row normalRow = sheet.createRow(lineCount);
                        Cell normalCell;

                        for (int cellCount = 0; cellCount < columnCount; cellCount++) {
                            normalCell = normalRow.createCell(cellCount);
                            CellStyle cellStyle = normalRow.getSheet().getWorkbook().createCellStyle();

                            try{
                                if(cellCount==0){
                                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                                    normalCell.setCellStyle(cellStyle);                                }
                                else{
                                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                                    normalCell.setCellStyle(cellStyle);
                                }
                                double number=Double.parseDouble(line.split(cvsSplitBy)[cellCount]);
                                normalCell.setCellValue(number);
                                normalCell.setCellStyle(normalStyle);
                                TotalColumnCount++;
                            }
                            catch(Exception e ){
                                if(cellCount==0){
                                    normalCell.getCellStyle().setAlignment(HorizontalAlignment.RIGHT);
                                }
                                else{
                                    normalCell.getCellStyle().setAlignment(HorizontalAlignment.LEFT);

                                }
                                normalCell.setCellValue(line.split(cvsSplitBy)[cellCount]);
                                normalCell.setCellStyle(normalStyle);
                                TotalColumnCount++;
                            }


                        }
                        lineCount++;


                    } else {

                        Row headerRow = sheet.createRow(lineCount);
                        Cell headerCell;

                        for (int cellCount = 0; cellCount < columnCount; cellCount++) {
                            headerCell = headerRow.createCell(cellCount);
                            headerCell.setCellValue(line.split(cvsSplitBy)[cellCount]);
                            headerCell.setCellStyle(headerStyle);
                            TotalColumnCount++;

                        }

                        headerflag = 1;
                        lineCount++;

                    }

                }

                // Autosize the Column to fit content
                for(int adjustSheet = 0; adjustSheet < TotalColumnCount; adjustSheet++) {
                    sheet.autoSizeColumn(adjustSheet);
                }


            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


            try {
                br.close();
            } catch (IOException e) {
                e.printStackTrace();
            }


        }

        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
