package com.example.myapp;

import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.text.TextUtils;
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

    // UI Elements
    private EditText idNumber, customerName, city, address, homeNumber, customerPhones, email,
            simNumber, mobileNumber, newNumber, creditCard, expiryMonth, expiryYear,
            cvv, notes;
    private RadioButton simDelivery, eSim;
    private CheckBox automaticVerification;
    private Button exportButton;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // Initialize UI elements
        idNumber = findViewById(R.id.id_number);
        customerName = findViewById(R.id.customer_name);
        city = findViewById(R.id.city);
        address = findViewById(R.id.address);
        homeNumber = findViewById(R.id.home_number);
        customerPhones = findViewById(R.id.customer_phones);
        email = findViewById(R.id.email);
        simNumber = findViewById(R.id.sim_number);
        mobileNumber = findViewById(R.id.mobile_number);
        creditCard = findViewById(R.id.credit_card);
        expiryMonth = findViewById(R.id.expiry_month);
        expiryYear = findViewById(R.id.expiry_year);
        cvv = findViewById(R.id.cvv);
        notes = findViewById(R.id.notes);
        simDelivery = findViewById(R.id.sim_delivery);
        eSim = findViewById(R.id.e_sim);
        automaticVerification = findViewById(R.id.automatic_verification);
        exportButton = findViewById(R.id.export_button);

        exportButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (areAllFieldsFilled()) {
                    exportDataToExcel();
                } else {
                    Toast.makeText(MainActivity.this, "אנא מלא את כל השדות", Toast.LENGTH_SHORT).show();
                }
            }
        });
    }

    private void exportDataToExcel() {
        // Create Excel workbook and sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Customer Data");

        // Add data to sheet
        addDataToSheet(sheet, getString(R.string.id_number), idNumber.getText().toString());
        addDataToSheet(sheet, getString(R.string.customer_name), customerName.getText().toString());
        addDataToSheet(sheet, getString(R.string.city), city.getText().toString());
        addDataToSheet(sheet, getString(R.string.address), address.getText().toString());
        addDataToSheet(sheet, getString(R.string.home_number), homeNumber.getText().toString());
        addDataToSheet(sheet, getString(R.string.customer_phones), customerPhones.getText().toString());
        addDataToSheet(sheet, getString(R.string.email), email.getText().toString());
        addDataToSheet(sheet, getString(R.string.sim_number), simNumber.getText().toString());
        addDataToSheet(sheet, getString(R.string.mobile_number), mobileNumber.getText().toString());
        addDataToSheet(sheet, getString(R.string.credit_card), creditCard.getText().toString());
        addDataToSheet(sheet, getString(R.string.expiry_month), expiryMonth.getText().toString());
        addDataToSheet(sheet, getString(R.string.expiry_year), expiryYear.getText().toString());
        addDataToSheet(sheet, getString(R.string.cvv), cvv.getText().toString());
        addDataToSheet(sheet, getString(R.string.notes), notes.getText().toString());
        // Add data for CheckBoxes and RadioButtons as "Yes/No"
        addDataToSheet(sheet, getString(R.string.sim_delivery), simDelivery.isChecked() ? "כן" : "לא");
        addDataToSheet(sheet, getString(R.string.e_sim), eSim.isChecked() ? "כן" : "לא");
        addDataToSheet(sheet, getString(R.string.automatic_verification), automaticVerification.isChecked() ? "כן" : "לא");

        try {
            // Generate filename using ONLY idNumber
            String safeCustomerId = idNumber.getText().toString().trim()
                    .replaceAll("[^0-9]", ""); // Remove non-numeric characters

            String fileName = safeCustomerId + ".xlsx";

            // Save the Excel file
            File file = new File(getFilesDir(), fileName);
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            // Notify user of success and offer to share
            Toast.makeText(this, getString(R.string.export_successful), Toast.LENGTH_SHORT).show();
            shareFile(this, file);

        } catch (IOException e) {
            e.printStackTrace();
            Toast.makeText(this, getString(R.string.export_failed), Toast.LENGTH_SHORT).show();
        }
    }

    // Helper function to add data to a new row in the sheet
    private void addDataToSheet(XSSFSheet sheet, String label, String value) {
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowNum + 1);
        Cell labelCell = row.createCell(0);
        labelCell.setCellValue(label);
        Cell valueCell = row.createCell(1);
        valueCell.setCellValue(value);
    }

    // Helper function to share the created file
    private void shareFile(Context context, File file) {
        Uri fileUri = FileProvider.getUriForFile(context,
                context.getPackageName() + ".provider", file);
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        shareIntent.putExtra(Intent.EXTRA_STREAM, fileUri);
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        context.startActivity(Intent.createChooser(shareIntent, "Share Excel File"));
    }

    // Function to check if all EditText fields are filled
    private boolean areAllFieldsFilled() {
        return !TextUtils.isEmpty(idNumber.getText().toString()) &&
                !TextUtils.isEmpty(customerName.getText().toString()) &&
                !TextUtils.isEmpty(city.getText().toString()) &&
                !TextUtils.isEmpty(address.getText().toString()) &&
                !TextUtils.isEmpty(homeNumber.getText().toString()) &&
                !TextUtils.isEmpty(customerPhones.getText().toString()) &&
                !TextUtils.isEmpty(email.getText().toString()) &&
                !TextUtils.isEmpty(simNumber.getText().toString()) &&
                !TextUtils.isEmpty(mobileNumber.getText().toString()) &&
                !TextUtils.isEmpty(creditCard.getText().toString()) &&
                !TextUtils.isEmpty(expiryMonth.getText().toString()) &&
                !TextUtils.isEmpty(expiryYear.getText().toString()) &&
                !TextUtils.isEmpty(cvv.getText().toString());
    }
}