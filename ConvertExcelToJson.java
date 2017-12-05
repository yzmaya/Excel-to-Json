/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package convert.excel.to.json;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;       

/**
 *
 * @author nestor.yzmaya
 */
public class ConvertExcelToJson {


    public static void main(String[] args) throws IOException {
        List sheetData = new ArrayList();

        FileInputStream fis = null;
try {
            //Se carga el archivo, aqui cambia la direccion para cargar tu archivo
            fis = new FileInputStream("C:\\Users\\nestor.yzmaya\\Documents\\CreaJson\\Hola.xlsx");
           
            //Se obtiene el libro de Excel
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            // Se obtiene la hoja del libro donde estan los datos
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator rows = sheet.rowIterator();
            //While para recorrer filas
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();

                Iterator cells = row.cellIterator();
                List data = new ArrayList();
                // While para recorrer celdas
                while (cells.hasNext()) {
                    XSSFCell cell = (XSSFCell) cells.next();
                    data.add(cell);
                }
                //Se agregan al objeto List
                sheetData.add(data);
            }
        } catch (IOException e) {
            System.out.println("Error IOException en leer archivo Service Implement: " + e.getMessage());
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
        // Termina try-catch-finally--------------------------------------------------------

         JSONObject innerObj = null;
        JSONObject  objFinal = new JSONObject();
        
        for (int i = 1; i < sheetData.size(); i++) {

            List list = (List) sheetData.get(i);
            innerObj = new JSONObject();

            for (int j = 0; j < list.size(); j++) {

                Cell cell = (Cell) list.get(j);
                
                //Se ponen nombres de las propiedades y se le asigna el valor dependiendo del indicador del arreglo
                if(j==1){innerObj.put("codigo", cell.toString());}
                if(j==2){innerObj.put("titulo", cell.toString());}
                if(j==3){innerObj.put("autor", cell.toString());}
                if(j==4){innerObj.put("idcoleccion", cell.toString());}
                if(j==5){innerObj.put("isbn", cell.toString());}
                if(j==6){innerObj.put("costo", cell.toString());}
                if(j==7){innerObj.put("existencia", cell.toString());}

            }
            // Se van agregando todos los objetos JSON que se van generando
            objFinal.put(i, innerObj);
        }
        
        try {
            // Se escribe el archivo .json
            FileWriter file = new FileWriter("C:\\Users\\nestor.yzmaya\\Documents\\CreaJson\\resultado.json");
           
            file.write(objFinal.toJSONString());
            file.flush();
            file.close();

        } catch (IOException e) {
            System.out.println("Error al generar archivo JSON: "+e.getMessage());
        }

        
    }
    
}
