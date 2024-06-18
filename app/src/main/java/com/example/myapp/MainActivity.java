package com.example.myapp;

import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.RadioButton;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    private EditText idNumber, customerName, city, address, homeNumber, customerPhones, email, simNumber, mobileNumber, newNumber, creditCard, expiryMonth, expiryYear, cvv, notes;
    private RadioButton simDelivery, eSim;
    private CheckBox automaticVerification;
    private Button exportButton; // You can add other buttons like add, remove, submit if needed

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // Initialize your views (EditText, RadioButton, CheckBox, Button)
        idNumber = findViewById(R.id.id_number);
        customerName = findViewById(R.id.customer_name);
        // ... Initialize other views

        exportButton = findViewById(R.id.export_button);

        exportButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                exportDataToExcel();
            }
        });
    }

    private void exportDataToExcel() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Customer Data");

        // Add data to Excel sheet
        addDataToSheet(sheet, 0, getString(R.string.id_number), idNumber.getText().toString());
        addDataToSheet(sheet, 1, getString(R.string.customer_name), customerName.getText().toString());
        // ... Add data for other fields similarly

        try {
            // Get the customer name and sanitize it for filename
            String safeCustomerName = customerName.getText().toString().trim().replaceAll("[^a-zA-Z0-9.-]", "_");
            if (safeCustomerName.isEmpty()) {
                safeCustomerName = "CustomerData";
            }
            String fileName = safeCustomerName + ".xlsx";

            File file = new File(getFilesDir(), fileName);
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            Toast.makeText(this, getString(R.string.export_successful), Toast.LENGTH_SHORT).show();

            shareFile(this, file);
        } catch (IOException e) {
            e.printStackTrace();
            Toast.makeText(this, getString(R.string.export_failed), Toast.LENGTH_SHORT).show();
        }
    }

    // Helper function to add data to Excel sheet
    private void addDataToSheet(XSSFSheet sheet, int rowNum, String label, String value) {
        Row row = sheet.createRow(rowNum);
        Cell labelCell = row.createCell(0);
        labelCell.setCellValue(label);
        Cell valueCell = row.createCell(1);
        valueCell.setCellValue(value);
    }

    private void shareFile(Context context, File file) {
        Uri fileUri = FileProvider.getUriForFile(context, context.getPackageName() + ".provider", file);
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        shareIntent.putExtra(Intent.EXTRA_STREAM, fileUri);
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

        context.startActivity(Intent.createChooser(shareIntent, "Share Excel File"));
    }
}