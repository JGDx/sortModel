import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class sortModel {
    static int count=0;
    public static void main(String[] args) {
        Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        List<Map<String,String>> list = null;
        String cellData = null;
        String filePath = "E:\\volvo.xls";
        String columns[] = {"title","model"};
        wb = readExcel(filePath);
        if(wb != null){
            //用来存放表中数据
            list = new ArrayList<Map<String,String>>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);

            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                Map<String,String> map = new LinkedHashMap<String,String>();
                row = sheet.getRow(i);
                if(row !=null){
//                    for (int j=0;j<colnum;j++){
//                        cellData = (String) getCellFormatValue(row.getCell(j));
//                        map.put(columns[j], cellData);
//                    }
                    String model=(String)getCellFormatValue(row.getCell(15));
                    if(model.equals("")||model==null){
//                        map.put(columns[0],(String)getCellFormatValue(row.getCell(0)));
//                        map.put(columns[1],(String)getCellFormatValue(row.getCell(15)));
//                        list.add(map);
                        model="Others";
                    }
                        String fileName=(String)getCellFormatValue(row.getCell(0));
                        String extension=fileName.substring(fileName.lastIndexOf('.')+1,fileName.length());
                        moveTotherFolders(model,fileName,extension);


                }else{
                    break;
                }

            }
        }
        System.out.println(count);
//        //遍历解析出来的list
//        for (Map<String,String> map : list) {
//            for (Entry<String,String> entry : map.entrySet()) {
//                System.out.print(entry.getKey()+":"+entry.getValue()+","+"count is:"+count);
//            }
//            System.out.println();
//        }

    }
    //读取excel
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else{
                return wb = null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }

    //移动文件
    public static void moveTotherFolders(String pathName,String fileName,String extension){
        try
        {
            String type="";
            if(extension.equals("tif")||extension.equals("ai")||
                    extension.equals("bmp")||extension.equals("eps")||
                    extension.equals("gif")||extension.equals("GIF")||
                    extension.equals("jpeg")||extension.equals("jpg")||
                    extension.equals("JPG")||extension.equals("psd")||
                    extension.equals("png")||extension.equals("PNG")||
                    extension.equals("TIF")){
                type="Images";
            }else if(extension.equals("flv")||extension.equals("mov")||
                    extension.equals("mp4")||extension.equals("wmv")||
                    extension.equals("mpg")||extension.equals("mpeg")){
                type="Films";
            }else if(extension.equals("dfont")||extension.equals("idml")||
                    extension.equals("indd")||extension.equals("otf")||extension.equals("srt")){
                type="Printing Material";
            }else if(extension.equals("pps")||extension.equals("pot")||
                    extension.equals("ppt")||extension.equals("pptx")){
                type="PPT";
            }else{
                type="Others";
            }
            if(pathName.contains("|")){
                pathName=pathName.substring(0,pathName.indexOf('|'));
            }
            fileName=fileName.substring(fileName.lastIndexOf('/')+1);
            File file=new File("G:\\Volvo Cars Content Store\\"+fileName); //源文件
//            if(!(file.exists())){
//                File folder=new File("G:\\Volvo Cars Content Store\\");
//                file=searchFile(folder,fileName);
//            }

            File path=new File("G:\\Volvo Cars Content Store\\"+pathName+"\\"+type+"\\");
//            File path=new File("G:\\Volvo Cars Content Store\\repeat");
            if(!path.exists()){
                path.mkdir();
            }
            File target=new File(path+"\\"+file.getName());
            if(file.exists()){
                if (file.renameTo(target)) //源文件移动至目标文件目录
                {

                    System.out.println("File is moved successful!");//输出移动成功
                }else
                {
                    System.out.println("File is failed to move !");//输出移动失败
                    System.out.println("file is :"+file+"target is :"+target);
                }
            }else{
                count++;
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    //查找文件集
    public static List<File> searchFiles(File folder,final String keyword){
        List<File> result=new ArrayList<File>();
        if(folder.isFile()){
            result.add(folder);
        }
        File[] subFolders=folder.listFiles(new FileFilter() {
            public boolean accept(File file) {
                if(file.isDirectory()){
                    return true;
                }
                if(file.getName().contains(keyword)){
                    return true;
                }
                return false;
            }
        });
        if(subFolders!=null){
            for(File file:subFolders){
                if(file.isFile()){
                    result.add(file);
                }else{
                    result.addAll(searchFiles(file,keyword));
                }
            }
        }
        return result;
    }
    //查找单个文件
    public static File searchFile(File folder,final String keyword){
        File result=new File(folder.getName()+"\\"+keyword);
        if(result.exists()){
            return result;
        }else{
            File[] subFolders=folder.listFiles(new FileFilter() {
                public boolean accept(File pathname) {
                    if(pathname.isDirectory()){
                        return true;
                    }
                    return false;
                }
            });
            if(subFolders!=null){
                for(File file:subFolders){
                    result= searchFile(file,keyword);
                    if(result.exists()){
                        break;
                    }
                }
            }
        }

        return result;
    }
}
