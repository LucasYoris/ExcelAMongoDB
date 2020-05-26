//--Nota:
/*ingresar los excel a convertir en "C:\cert\Migración datos MongoDB\MigracionExcelMongoDB\Excels entrada\"
verificar que no hayan columnas vacias sino puede tirar NullException
los archivos json para MongoDB quedaran en la carpeta "C:\cert\Migración datos MongoDB\MigracionExcelMongoDB\Json Salida"*/
package replacejsoncharacters;

import java.io.IOException;
/**
 *
 * @author lyoris
 */
public class CrearJSONFromXLSX {    
    
    public static void main(String[] args) throws IOException {
        ReadingExcel leer = new ReadingExcel();
        leer.read();
    }
}