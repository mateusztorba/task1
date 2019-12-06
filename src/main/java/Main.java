import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {

    private static final String XLSX_FILE_NAME = "test1.xlsx";
    private static final int NUMBER_COLUMNS = 3;
    private static final int ID_POSITION_COLUMN = 3;

    public static void main(String[] args) throws IOException {
        print(getObjectAsJson(processXmlFile(XLSX_FILE_NAME)));
    }

    private static void print(String objectAsJson) {
        System.out.println(objectAsJson);
    }

    private static String getObjectAsJson(List<Node> nodes) throws JsonProcessingException {
        return new ObjectMapper().writeValueAsString(nodes);
    }

    private static List<Node> processXmlFile(String filePath) throws IOException {
        List<Node> nodes = new ArrayList<Node>();
        Iterator<Row> rowIterator = new XSSFWorkbook(new FileInputStream(new File(filePath))).getSheetAt(0).iterator();
        Node currentNode = null;
        int currentNodePosition = 1;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if(row.getRowNum() == 0) {
                continue;
            }
            for (int currentColumnIndex = 0; currentColumnIndex < NUMBER_COLUMNS; currentColumnIndex++) {
                Cell cell = row.getCell(currentColumnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    if (currentColumnIndex == 0) {
                        currentNode = new Node((int) (row.getCell(ID_POSITION_COLUMN).getNumericCellValue()), cell.getStringCellValue());
                        nodes.add(currentNode);
                        currentNodePosition = 1;
                    } else {
                        if (currentColumnIndex == currentNodePosition) {
                            currentNodePosition = currentColumnIndex;
                            currentNode.getNodes().add(new Node((int) (row.getCell(ID_POSITION_COLUMN).getNumericCellValue()), cell.getStringCellValue()));
                        } else {
                            currentNode.getNodes().get(currentNode.getNodes().size() - 1).getNodes().add(new Node((int) (row.getCell(ID_POSITION_COLUMN).getNumericCellValue()), cell.getStringCellValue()));
                        }
                    }
                }
            }
        }
        return nodes;
    }
}
