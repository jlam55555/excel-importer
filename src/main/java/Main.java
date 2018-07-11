// apache POI
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

// file io
import java.io.File;
import java.io.IOException;

// json
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;

// google maps
import com.google.maps.*;
import com.google.maps.errors.ApiException;
import com.google.maps.model.LatLng;

// util classes
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

public class Main {

  // file path
  public static final String FILE_PATH = "./in/doctors.xlsx";

  public static void main(String[] args)
    throws IOException, InvalidFormatException, InterruptedException, ApiException {

    // get first sheet of spreadsheet and ignore errors
    Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));
    Sheet sheet = workbook.getSheetAt(0);
    System.out.printf("%nThe errors above are from the Apache POI and are well-documented to be safely ignored.%n");

    // create data formatter to interpret data
    DataFormatter df = new DataFormatter();

    // iterate over rows and columns
    List<Map<String, Object>> data = new ArrayList<>();
    List<String> headers = new ArrayList();
    AtomicBoolean headersSet = new AtomicBoolean(false);

    // create google maps api object
    GeoApiContext gmaps = new GeoApiContext.Builder()
        .apiKey("AIzaSyCb1594vVZ1GlBBHkqOTHBVKjYsZOKAas4")
        .queryRateLimit(40)
        .build(); // make sure it fits within the free 50qps limit

    sheet.forEach(row -> {
      // set headers on first run
      if(!headersSet.getAndSet(true)) {
        row.forEach(cell -> headers.add(df.formatCellValue(cell)));

        // VALIDATION: check that all necessary rows are created (capitalization counts!)
        // Name, Gender, Specialty, VIP, Address, Language
        if(headers.indexOf("Name") == -1 || headers.indexOf("Gender") == -1 || headers.indexOf("Specialty") == -1 || headers.indexOf("VIP") == -1 || headers.indexOf("Address") == -1 || headers.indexOf("Language") == -1) {
          System.err.printf("Error: all of the following headers: \"Name, Gender, Specialty, Language, VIP, Address\" must be present (capitalization matters, but order does not).%n");
        }

      }
      // add doctor data on subsequent runs
      else {
        Map<String, Object> rowData = new HashMap<>();
        AtomicInteger index = new AtomicInteger(0);
        row.forEach(cell -> {
          String value = df.formatCellValue(cell);
          String header = headers.get(index.getAndIncrement());

          // FORMATTING: lowercase gender, vip, specialty, language
          if(header.equals("Gender") || header.equals("Specialty") || header.equals("VIP") || header.equals("Language")) {
            value = value.toLowerCase();
          }

          // VALIDATION: checking values for gender, vip
          if(header.equals("Gender") && !value.equals("male") && !value.equals("female")) {
            System.err.printf("Error: \"Gender\" field must only have values \"male\" or \"female\" for doctor \"%s\".%n", rowData.get("Name"));
          }
          if(header.equals("VIP") && !value.equals("yes") && !value.equals("no")) {
            System.err.printf("Error: \"VIP\" field must only have values \"yes\" or \"no\" for doctor \"%s\".%n", rowData.get("Name"));
          }

          rowData.put(header, value);
        });

        // VALIDATION: check that place exists
        // get place coordinates
        try {
          LatLng location = GeocodingApi.geocode(gmaps, (String) rowData.get("Address")).await()[0].geometry.location;
          Map<String, Double> locationMap = new HashMap<>();
          locationMap.put("lat", location.lat);
          locationMap.put("lng", location.lng);
          rowData.put("coords", locationMap);
        } catch(Exception e) {
          System.err.printf("Error: No address found for doctor \"%s\" with address \"%s\". Make sure address exists on Google Maps.%n", rowData.get("Name"), rowData.get("Address"));
        }

        data.add(rowData);
      }
    });

    // write to JSON
    ObjectMapper mapper = new ObjectMapper();
    mapper.enable(SerializationFeature.INDENT_OUTPUT);

    // write JSON file
    String outputPath = "./out/doctors.json";
    File file = new File(outputPath);
    mapper.writeValue(file, data);

    // write backup
    String backupOutputPath = "./out/doctors-" + new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(new Date()) + ".json";
    File backupFile = new File(backupOutputPath);
    mapper.writeValue(backupFile, data);

    outputPath = outputPath.substring(1);
    backupOutputPath = backupOutputPath.substring(1);

    // print instructions
    System.out.printf("Import from Excel to JSON complete.%nOutput JSON has been written out to path \"%s%s\".%nA backup file has been written to \"%s%s\"%nCopy the output JSON file to the myGUT directory.%n", System.getProperty("user.dir"), outputPath, System.getProperty("user.dir"), backupOutputPath);

    // close context
    gmaps.shutdown();

  }

}
