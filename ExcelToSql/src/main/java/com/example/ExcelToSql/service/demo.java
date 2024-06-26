package com.example.ExcelToSql.service;

public class demo {
    public List<List<String>> readFile(MultipartFile file) throws IOException {
        Workbook workbook = null;
        InputStream inputStream=null;
        List<List<String>> data = new ArrayList<>();
        try {
            inputStream = file.getInputStream();
            workbook = new XSSFWorkbook(inputStream);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            rows.next();
            while (rows.hasNext()) {
                Row currentRow = rows.next();
                List<String> rowData = new ArrayList<>();
                for(int col=0; col< currentRow.getLastCellNum();col++){
                    Cell column = currentRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if(column!=null){
                        switch (column.getCellType()) {
                            case STRING:
                                rowData.add(column.getStringCellValue());
                                break;
                            case NUMERIC:
                                rowData.add(String.valueOf(column.getStringCellValue()));
                                break;
                            default:
                                rowData.add("NULL");
                                break;
                        }
                    }
                    else{
                        rowData.add("NULL");
                    }

                }

                data.add(rowData);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        finally{
            if(workbook!=null){
                workbook.close();
            }
            if(inputStream!=null){
                inputStream.close();
            }
        }
        String fileName=file.getOriginalFilename();
        assert fileName != null;
        file.transferTo(new File(RFI_REGISTER_VOLUME_PATH ,fileName));
        purgeFilesOnLastModified();
        return data;
    }

}
