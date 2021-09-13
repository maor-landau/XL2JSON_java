package reader;
import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.Writer;

public class ReadExcel {
    private String inputFile;
    public void setInputFile(String inputFile) {
        this.inputFile = inputFile;
    }
    public void read() throws IOException {
        File inputWorkbook = new File(inputFile);
        Workbook w;
        Writer output;
        output = new BufferedWriter(new FileWriter("ROOT:\\Users\\{{YOURUSERNAME}}\\Downloads\\{{FILENAME}}"));  //clears file every time
        try {
            w = Workbook.getWorkbook(inputWorkbook);
            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            // Loop over first 10 column and lines
            output.append("{");
            for (int j = 0; j < sheet.getColumns(); j++) {
                for (int i = 0; i < sheet.getRows(); i++) {
                    Cell cell1 = sheet.getCell(0, i);
                    Cell cell2 = sheet.getCell(1, i);
                    output.append(cell1.getContents()+":"+cell2.getContents() + ",");
                    System.out.println("{'name' : "+"'"+ cell1.getContents()+" : "+cell2.getContents()+"'"+"},");
                }
            }
            output.append("}");
            output.close();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        ReadExcel test = new ReadExcel();
        String fileDir = "C:\\Users\\{{YOURUSERNAME}}\\Downloads\\{{FILENAME}}";
        test.setInputFile(fileDir);
        test.read();
    }
}