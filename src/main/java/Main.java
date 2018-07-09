// apache POI
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

// file io
import java.io.File;
import java.io.IOException;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

// json
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;

// util classes
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

  // file path
  public static final String FILE_PATH = "./in/doctors.xlsx";

  public static void main(String[] args)
    throws IOException, InvalidFormatException {

    // get first sheet of spreadsheet
    Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));
    Sheet sheet = workbook.getSheetAt(0);

    // create data formatter to interpret data
    DataFormatter df = new DataFormatter();

    // iterate over rows and columns
    List<Map<String, String>> data = new ArrayList<>();
    List<String> headers = new ArrayList();
    AtomicBoolean headersSet = new AtomicBoolean(false);

    sheet.forEach(row -> {
      // set headers on first run
      if(!headersSet.getAndSet(true)) {
        row.forEach(cell -> headers.add(df.formatCellValue(cell)));
      }
      // add doctor data on subsequent runs
      else {
        Map<String, String> rowData = new HashMap<>();
        AtomicInteger index = new AtomicInteger(0);
        row.forEach(cell -> rowData.put(headers.get(index.getAndIncrement()), df.formatCellValue(cell)));
        data.add(rowData);
      }
    });

    // write to JSON
    ObjectMapper mapper = new ObjectMapper();
    mapper.enable(SerializationFeature.INDENT_OUTPUT);
    File file = new File("./out/doctors.json");
    mapper.writeValue(file, data);

  }

}
