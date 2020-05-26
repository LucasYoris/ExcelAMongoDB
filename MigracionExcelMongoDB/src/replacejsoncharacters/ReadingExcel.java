package replacejsoncharacters;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author lyoris
 */
public class ReadingExcel {
    
    public void read() throws IOException{
        
        //try {
        //se toma la cantidad de excels en la carpeta Excels entrada
        File archivosEntrada = new File("C:\\cert\\Desarrollos para automatización\\MigracionExcelMongoDB\\Excels entrada\\");
        File[] listOfFiles = archivosEntrada.listFiles();
        //se lee cada excel
        for (File file : listOfFiles) {
        if (file.isFile()) {
            //archivo json
            file.getName();
            String pathSalida = "C:/cert/Desarrollos para automatización/MigracionExcelMongoDB/Json Salida/"+file.getName();
            pathSalida=pathSalida.replaceAll(".xlsx", ".json");
            FileWriter json = new FileWriter(pathSalida);
            //archivo excel
            FileInputStream archivoXlsx = new FileInputStream(new File(file.getPath()));
            XSSFWorkbook wb = new XSSFWorkbook(archivoXlsx);
            //se toma la cantidad de hojas en el wb
            int sheets = wb.getNumberOfSheets();
            //se lee cada hoja del excel
            for(int i=0;i<sheets;i++){
                XSSFSheet sheet = wb.getSheetAt(i);
                //nombre de la hoja
                String sheetName = sheet.getSheetName();
                //cantidad de filas
                int rowCount = sheet.getPhysicalNumberOfRows();
                //cantidad de columnas
                int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
                
                //tomar header de columnas
                List<String> header = new ArrayList();
                Row fila0 = sheet.getRow(0);
                for(int j=0;j<columnCount;j++){
                    if(fila0.getCell(j).toString().equals("Denominacion")){
                        header.add(j, fila0.getCell(j).toString().toLowerCase());
                    }else{
                        header.add(j, fila0.getCell(j).toString());
                    }
                }
                
                String cadena="";
                
                for(int a=1;a<rowCount;a++){
                    Row fila = sheet.getRow(a);
                    cadena = "{\"identificador\":"+"\""+sheet.getSheetName()+"\""+",";
                    for(int b=0;b<columnCount;b++){
                        Cell celda = fila.getCell(b);
                        switch(celda.getCellType().toString()){
                            case "NUMERIC":
                                if(b!=(columnCount-1)){
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+((int)celda.getNumericCellValue())+"\""+",";
                                }else{
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+((int)celda.getNumericCellValue())+"\""+"}";
                                    json.write(cadena+"\n");
                                }
                                if(b==(columnCount-1)){
                                    System.out.println(cadena);}
                                break;
                            case "STRING":
                                if(b!=(columnCount-1)){
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getStringCellValue()+"\""+",";
                                }else{
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getStringCellValue()+"\""+"}";
                                    json.write(cadena+"\n");
                                }
                                if(b==(columnCount-1)){
                                    System.out.println(cadena);}
                                break;
                            case "FORMULA":
                                if(b!=(columnCount-1)){
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+",";
                                }else{
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+"}";
                                    json.write(cadena+"\n");
                                }
                                if(b==(columnCount-1)){
                                    System.out.println(cadena);}
                                break;
                            case "BOOLEAN":
                                if(b!=(columnCount-1)){
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+",";
                                }else{
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+"}";
                                    json.write(cadena+"\n");
                                }
                                if(b==(columnCount-1)){
                                    System.out.println(cadena);}
                                break;
                            case "ERROR":
                                if(b!=(columnCount-1)){
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+",";
                                }else{
                                    cadena = cadena+"\""+header.get(b)+"\""+":"+"\""+celda.getCellFormula()+"\""+"}";
                                    json.write(cadena+"\n");
                                if(b==(columnCount-1)){
                                    System.out.println(cadena);}
                                break;
                        }
                    }                  
                }
                System.out.println("Sheet "+sheet.getSheetName()+" finalizada");
        }   
            }
            json.close();
        }
        }
    }
}
    
      
