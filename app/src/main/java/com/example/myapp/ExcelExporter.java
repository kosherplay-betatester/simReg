import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Environment;
import android.widget.EditText;
import android.widget.Toast;

import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelExporter {

    public static void exportToExcel(Context context, EditText... editTexts) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("User Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < editTexts.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue("Field " + (i + 1));
        }

        // Create data row
        Row dataRow = sheet.createRow(1);
        for (int i = 0; i < editTexts.length; i++) {
            Cell cell = dataRow.createCell(i);
            cell.setCellValue(editTexts[i].getText().toString());
        }

        // Save Excel file
        String fileName = "UserData.xlsx";
        File filePath = new File(context.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS), fileName);
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            Toast.makeText(context, "Excel file saved to Downloads", Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            e.printStackTrace();
            Toast.makeText(context, "Error saving Excel file", Toast.LENGTH_SHORT).show();
        }

        // Share Excel file
        shareFile(context, filePath);
    }

    private static void shareFile(Context context, File file) {
        Uri fileUri = FileProvider.getUriForFile(context, context.getPackageName() + ".provider", file);
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        shareIntent.putExtra(Intent.EXTRA_STREAM, fileUri);
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

        context.startActivity(Intent.createChooser(shareIntent, "Share Excel File"));
    }
}
