   public static String retrieve1(String Col)
        {
            String rowval = null;
           
            XSSFSheet sh;
            XSSFWorkbook wb;
            string path = CommonVariable.TESTDATA.Trim() + "Data.xlsx";
            if (File.Exists(path))
            {
                wb = new XSSFWorkbook(path);
                // create sheet
                sh = (XSSFSheet)wb.GetSheet("Sheet1");
                // 3 rows, 2 columns
              

        //         Iterator<Row> rowIterator = sheet.iterator();
        //while(rowIterator.hasNext()) {
        //    Row row = rowIterator.next();
        //    Iterator<Cell> cellIterator = row.cellIterator();
        //    while(cellIterator.hasNext()) {
        //        Cell cell = cellIterator.next();





                int r = sh.LastRowNum;
                for (int j = 0; j <=r; j++)
                {
                  XSSFRow row = (XSSFRow)sh.GetRow(j);
                  int c = row.LastCellNum;
                    for (int i = 0; i < c; i++)
                    {
                        rowval = row.GetCell(i).StringCellValue.ToString();

                    }
                }
            }

            return rowval;
        }
