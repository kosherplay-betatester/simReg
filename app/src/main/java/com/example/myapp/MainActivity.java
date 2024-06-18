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
    private Button addButton, removeButton, submitButton, exportButton;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        idNumber = findViewById(R.id.id_number);
        customerName = findViewById(R.id.customer_name);
        city = findViewById(R.id.city);
        address = findViewById(R.id.address);
        homeNumber = findViewById(R.id.home_number);
        customerPhones = findViewById(R.id.customer_phones);
        email = findViewById(R.id.email);
        simNumber = findViewById(R.id.sim_number);
        simDelivery = findViewById(R.id.sim_delivery);
        eSim = findViewById(R.id.e_sim);
        mobileNumber = findViewById(R.id.mobile_number);
        newNumber = findViewById(R.id.new_number);
        automaticVerification = findViewById(R.id.automatic_verification);
        creditCard = findViewById(R.id.credit_card);
        expiryMonth = findViewById(R.id.expiry_month);
        expiryYear = findViewById(R.id.expiry_year);
        cvv = findViewById(R.id.cvv);
        notes = findViewById(R.id.notes);

        addButton = findViewById(R.id.add_button);
        removeButton = findViewById(R.id.remove_button);
        submitButton = findViewById(R.id.submit_button);
        exportButton = findViewById(R.id.export_button);

        submitButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                Toast.makeText(MainActivity.this, getString(R.string.form_submitted), Toast.LENGTH_SHORT).show();
            }
        });

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

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(getString(R.string.id_number));
        cell = row.createCell(1);
        cell.setCellValue(idNumber.getText().toString());

        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.customer_name));
        cell = row.createCell(1);
        cell.setCellValue(customerName.getText().toString());

        row = sheet.createRow(2);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.city));
        cell = row.createCell(1);
        cell.setCellValue(city.getText().toString());

        row = sheet.createRow(3);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.address));
        cell = row.createCell(1);
        cell.setCellValue(address.getText().toString());

        row = sheet.createRow(4);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.home_number));
        cell = row.createCell(1);
        cell.setCellValue(homeNumber.getText().toString());

        row = sheet.createRow(5);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.customer_phones));
        cell = row.createCell(1);
        cell.setCellValue(customerPhones.getText().toString());

        row = sheet.createRow(6);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.email));
        cell = row.createCell(1);
        cell.setCellValue(email.getText().toString());

        row = sheet.createRow(7);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.sim_number));
        cell = row.createCell(1);
        cell.setCellValue(simNumber.getText().toString());

        row = sheet.createRow(8);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.sim_delivery));
        cell = row.createCell(1);
        cell.setCellValue(simDelivery.isChecked() ? "Yes" : "No");

        row = sheet.createRow(9);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.e_sim));
        cell = row.createCell(1);
        cell.setCellValue(eSim.isChecked() ? "Yes" : "No");

        row = sheet.createRow(10);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.mobile_number));
        cell = row.createCell(1);
        cell.setCellValue(mobileNumber.getText().toString());

        row = sheet.createRow(11);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.new_number));
        cell = row.createCell(1);
        cell.setCellValue(newNumber.getText().toString());

        row = sheet.createRow(12);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.automatic_verification));
        cell = row.createCell(1);
        cell.setCellValue(automaticVerification.isChecked() ? "Yes" : "No");

        row = sheet.createRow(13);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.credit_card));
        cell = row.createCell(1);
        cell.setCellValue(creditCard.getText().toString());

        row = sheet.createRow(14);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.expiry_month));
        cell = row.createCell(1);
        cell.setCellValue(expiryMonth.getText().toString());

        row = sheet.createRow(15);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.expiry_year));
        cell = row.createCell(1);
        cell.setCellValue(expiryYear.getText().toString());

        row = sheet.createRow(16);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.cvv));
        cell = row.createCell(1);
        cell.setCellValue(cvv.getText().toString());

        row = sheet.createRow(17);
        cell = row.createCell(0);
        cell.setCellValue(getString(R.string.notes));
        cell = row.createCell(1);
        cell.setCellValue(notes.getText().toString());

        try {
            File file = new File(getFilesDir(), "CustomerData.xlsx");
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

    private void shareFile(Context context, File file) {
        Uri fileUri = FileProvider.getUriForFile(context, context.getPackageName() + ".provider", file);
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        shareIntent.putExtra(Intent.EXTRA_STREAM, fileUri);
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

        context.startActivity(Intent.createChooser(shareIntent, "Share Excel File"));
    }
}
