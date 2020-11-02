## 对excel进行操作的demo项目
- POI 使用POI进行操作（03版和07版的excel）
- EasyExcel 使用阿里提供的EasyExcel框架进行操作

````java
public static void main(String[] args) throws Exception{
            // 创建工作表
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 创建sheet页
            Sheet sheet = workbook.createSheet("测试03版本");
            // 创建第一行数据
            Row row = sheet.createRow(0);
            // 第一行创建10个cell并插入数据
            for (int i = 0; i < 10; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue("测试03第一行cell："+i);
            }

            Row row1 = sheet.createRow(1);
            for (int i = 0; i < 10; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue("测试03第二行cell："+i);
            }

            // 03版本的excel为xsl格式文件
            FileOutputStream fileOutputStream = new FileOutputStream("test03.xsl");

            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

    }
````