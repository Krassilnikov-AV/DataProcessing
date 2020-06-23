
package poi.main;

import java.io.FileNotFoundException;
import java.io.IOException;

import poi.createFiles.Create;

public class Main {

     public static void main(String[] args) throws FileNotFoundException, IOException {
     
         Create cf = new Create();
         cf.create();
    }
    
}
