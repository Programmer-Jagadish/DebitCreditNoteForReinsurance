package com.reinsurance.notes;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class BrokerDebitCreditGenerator {

    public static void main(String[] args) {

        // ✅ Detect base directory (works for both IntelliJ and JAR/BAT)
        String basePath = System.getProperty("user.dir");

        // Default locations for packaged use
        String excelFilePath = basePath + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
        String templatePath = basePath + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
        String creditTemplatePath = basePath + File.separator + "resources" + File.separator + "CreditNoteTemplate.docx"; // NEW
        String outputFolder = basePath + File.separator + "resources" + File.separator + "output" + File.separator;

        // ✅ Fallback for IntelliJ execution (src/main/resources)
        File excelFile = new File(excelFilePath);
        if (!excelFile.exists()) {
            excelFilePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
            templatePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
            creditTemplatePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "CreditNoteTemplate.docx";
            outputFolder = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "output" + File.separator;
        }

        System.out.println("==============================================");
        System.out.println("   Reinsurance Debit & Credit Note Generator");
        System.out.println("==============================================");
        System.out.println("Excel File: " + excelFilePath);
        System.out.println("Debit Template: " + templatePath);
        System.out.println("Credit Template: " + creditTemplatePath);
        System.out.println("Output Folder: " + outputFolder);
        System.out.println("Processing data...\n");

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet mainSheet = wb.getSheetAt(0);
            Sheet creditSheet = wb.getSheet("CreditNoteDetails"); // may be null if not present

            File outDir = new File(outputFolder);
            if (!outDir.exists()) outDir.mkdirs();

            DateTimeFormatter df = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
            int mainProcessedCol = 21; // existing main sheet processed column index

            for (int r = 1; r <= mainSheet.getLastRowNum(); r++) {
                Row row = mainSheet.getRow(r);
                if (row == null || isRowEmpty(row)) {
                    System.out.println("⚠️ Skipping blank row " + r);
                    continue;
                }

                String processedFlag = getString(row, mainProcessedCol);
                if (processedFlag.equalsIgnoreCase("Yes") || processedFlag.equalsIgnoreCase("Processed")) {
                    System.out.println("⏩ Skipping main row " + r + " (already processed)");
                    continue;
                }

                // --- Input Columns from main sheet ---
                String debitNoteNo = getString(row, 0);
                if (debitNoteNo == null || debitNoteNo.isEmpty()) {
                    debitNoteNo = "DN-" + String.format("%03d", r);
                }

                String docDate = getString(row, 1);
                if (docDate.isEmpty()) docDate = LocalDate.now().format(df);

                String interest = getString(row, 2);
                String insured = getString(row, 3);
                String defaultReinsurer = getString(row, 4);
                String period = getString(row, 5);

                double SI = getDouble(row, 6);
                double cedentRate = getDouble(row, 7);
                double mainReinsRate = getDouble(row, 8);
                double share = getDouble(row, 9);
                double brokerage = getDouble(row, 10);
                double cedingCommPct = getDouble(row, 15); // input ceding commission % in main

                if (SI == 0 || cedentRate == 0 || share == 0) {
                    System.out.println("⚠️ Skipping incomplete main row " + r);
                    continue;
                }

                // --- Calculations for debit + reinsurer summary ---
                // Cedent side
                double grossPremiumCedent = SI * (cedentRate / 100);
                double sharePremiumCedent = grossPremiumCedent * (share / 100);

                // Reinsurer side (summary using mainReinsRate if present, else cedentRate)
                double effectiveReinsRateMain = (mainReinsRate > 0) ? mainReinsRate : cedentRate;
                double grossPremiumReinsMain = SI * (effectiveReinsRateMain / 100);
                double sharePremiumReinsMain = grossPremiumReinsMain * (share / 100);

                // Common deductions (summary)
                double cedingCommissionAmtMain = sharePremiumCedent * (cedingCommPct / 100);
                double grossBrokerageMain = sharePremiumReinsMain * (brokerage / 100);
                double netBrokerageMain = grossBrokerageMain / 2.0;

                // Final amounts summary
                double netPremiumFromYou = sharePremiumCedent - cedingCommissionAmtMain; // debit
                double netPremiumToYou = sharePremiumReinsMain - netBrokerageMain - cedingCommissionAmtMain; // credit summary

                // --- Write summary back to main sheet (existing indices, unchanged) ---
                setNumeric(row, 11, grossPremiumCedent);
                setNumeric(row, 12, sharePremiumCedent);
                setNumeric(row, 13, grossPremiumReinsMain);
                setNumeric(row, 14, sharePremiumReinsMain);
                setNumeric(row, 15, cedingCommPct);
                setNumeric(row, 16, cedingCommissionAmtMain);
                setNumeric(row, 17, grossBrokerageMain);
                setNumeric(row, 18, netBrokerageMain);
                setNumeric(row, 19, netPremiumFromYou);
                setNumeric(row, 20, netPremiumToYou);
                setString(row, mainProcessedCol, "Yes"); // mark main processed

                // --- Generate Debit Note Word (same as before) ---
                String safeFileName = debitNoteNo.replaceAll("[^a-zA-Z0-9-_]", "_");
                generateDebitNote(
                        templatePath,
                        outputFolder + safeFileName + ".docx",
                        debitNoteNo,
                        docDate,
                        interest,
                        insured,
                        defaultReinsurer,
                        period,
                        SI,
                        cedentRate,
                        grossPremiumCedent,
                        share,
                        sharePremiumCedent,
                        netPremiumFromYou
                );

                System.out.println("✅ Main Debit Note generated: " + debitNoteNo);

                // --- NEW: Generate Credit Notes for each matching row in CreditNoteDetails ---
                if (creditSheet != null) {
                    for (int cr = 1; cr <= creditSheet.getLastRowNum(); cr++) {
                        Row crow = creditSheet.getRow(cr);
                        if (crow == null) continue;

                        String linkedDebit = getString(crow, 0);
                        if (!linkedDebit.equalsIgnoreCase(debitNoteNo)) continue; // not for this debit

                        String creditProcessed = getString(crow, 6);
                        if (creditProcessed.equalsIgnoreCase("Yes") || creditProcessed.equalsIgnoreCase("Processed")) {
                            System.out.println("⏩ Skipping credit row " + cr + " for " + linkedDebit + " (already processed)");
                            continue;
                        }

                        // read credit row inputs
                        String reinsurerName = getString(crow, 1);
                        double reinsurerShare = getDouble(crow, 2); // percent for that reinsurer
                        double creditRowRate = getDouble(crow, 3); // optional rate per credit row
                        double creditRowBrokerage = getDouble(crow, 4); // optional brokerage override
                        double creditRowCedingPct = getDouble(crow, 5); // optional ceding commission%

                        // choose effective values (priority: credit row -> main sheet rates)
                        double effectiveReinsRate = (creditRowRate > 0) ? creditRowRate : ((mainReinsRate > 0) ? mainReinsRate : cedentRate);
                        double effectiveBrokerage = (creditRowBrokerage > 0) ? creditRowBrokerage : brokerage;
                        double effectiveCedingPct = (creditRowCedingPct > 0) ? creditRowCedingPct : cedingCommPct;

                        // calculations per reinsurer
                        double grossPremiumForReinsurer = SI * (effectiveReinsRate / 100);
                        double sharePremiumForReinsurer = grossPremiumForReinsurer * (reinsurerShare / 100);
                        double commissionAmtForReinsurer = sharePremiumForReinsurer * (effectiveCedingPct / 100);
                        double grossBrokerageForReinsurer = sharePremiumForReinsurer * (effectiveBrokerage / 100);
                        double netBrokerageForReinsurer = grossBrokerageForReinsurer / 2.0;
                        double netPayableToReinsurer = sharePremiumForReinsurer - netBrokerageForReinsurer - commissionAmtForReinsurer;

                        // write some calculated outputs back to credit sheet for visibility (cols 7..11)
                        setNumeric(crow, 7, grossPremiumForReinsurer);       // Fac Premium 100%
                        setNumeric(crow, 8, sharePremiumForReinsurer);       // Share Premium
                        setNumeric(crow, 9, commissionAmtForReinsurer);      // Ceding Commission Amount
                        setNumeric(crow, 10, grossBrokerageForReinsurer);    // Gross Brokerage
                        setNumeric(crow, 11, netPayableToReinsurer);         // Net Payable

                        // generate credit note doc
                        String safeCreditName = ("CreditNote_" + debitNoteNo + "_" + reinsurerName).replaceAll("[^a-zA-Z0-9-_\\.]", "_");
                        try {
                            generateCreditNote(
                                    creditTemplatePath,
                                    outputFolder + safeCreditName + ".docx",
                                    // template fields
                                    "CN-" + safeCreditName,                     // Credit Note No (basic unique id)
                                    LocalDate.now().format(df),                // Document Date
                                    interest,
                                    insured,
                                    reinsurerName,
                                    period,
                                    SI,
                                    effectiveReinsRate,
                                    grossPremiumForReinsurer,
                                    reinsurerShare,
                                    sharePremiumForReinsurer,
                                    grossBrokerageForReinsurer,
                                    netPayableToReinsurer
                            );
                            // mark credit row as processed
                            setString(crow, 6, "Yes");
                            System.out.println("   ✅ Credit Note generated for " + reinsurerName + " (linked to " + debitNoteNo + ")");
                        } catch (Exception ce) {
                            System.err.println("   ❌ Failed to generate credit note for " + reinsurerName + ": " + ce.getMessage());
                            ce.printStackTrace();
                        }
                    } // end credit sheet loop
                } // if creditSheet != null

            } // end main sheet loop

            // Save updated Excel (main + credit sheet writes)
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                wb.write(fos);
            }

            System.out.println("\n✅ All Debit & Credit Notes Processed and Excel Updated Successfully.");
            openOutputFolder(outputFolder);

        } catch (FileNotFoundException e) {
            System.err.println("❌ File not found! Please check folder structure:");
            System.err.println("Expected Excel file at: " + excelFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // --- Utility: Check if row is empty ---
    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int c = 0; c <= 10; c++) {
            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty())
                    return false;
                if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() != 0)
                    return false;
            }
        }
        return true;
    }

    // --- Excel Helpers ---
    private static String getString(Row row, int idx) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return "";
        if (c.getCellType() == CellType.STRING) return c.getStringCellValue().trim();
        if (c.getCellType() == CellType.NUMERIC) return String.valueOf(c.getNumericCellValue());
        return "";
    }

    private static double getDouble(Row row, int idx) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return 0.0;
        if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();
        if (c.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(c.getStringCellValue().replace(",", "").trim());
            } catch (Exception e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private static void setNumeric(Row row, int idx, double val) {
        if (idx > 50) return; // safe guard — we only use small indices
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.NUMERIC);
        c.setCellValue(val);
    }

    private static void setString(Row row, int idx, String val) {
        if (idx > 50) return;
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.STRING);
        c.setCellValue(val);
    }

    // --- Word Template Generator for Debit Note (existing) ---
    private static void setCellText(XWPFTable table, int rowIdx, int colIdx, String text) {
        XWPFTableRow row = table.getRow(rowIdx);
        if (row != null && row.getCell(colIdx) != null) {
            row.getCell(colIdx).removeParagraph(0);
            row.getCell(colIdx).setText(text);
        }
    }

    private static void generateDebitNote(
            String templatePath, String outputPath,
            String debitNoteNo, String documentDate, String interest,
            String insured, String reinsurer, String period,
            double sumInsured, double rate, double facPremiumFull,
            double share, double sharePremium, double netPremiumFromYou)
            throws IOException, InvalidFormatException {

        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            XWPFTable table = doc.getTables().get(0);

            setCellText(table, 0, 2, debitNoteNo);
            setCellText(table, 1, 2, documentDate);
            setCellText(table, 3, 2, interest);
            setCellText(table, 4, 2, insured);
            setCellText(table, 5, 2, reinsurer);
            setCellText(table, 6, 2, period);
            setCellText(table, 7, 2, "USD " + format(sumInsured));
            setCellText(table, 8, 2, String.format("%.2f%%", rate));
            setCellText(table, 9, 2, "USD " + format(facPremiumFull));
            setCellText(table, 10, 2, String.format("%.2f%% of 100%%", share));
            setCellText(table, 11, 2, "USD " + format(sharePremium));
            setCellText(table, 12, 2, "USD " + format(netPremiumFromYou));

            File outFile = new File(outputPath);
            outFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                doc.write(fos);
            }
        }
    }

    // --- Word Template Generator for Credit Note (NEW) ---
    private static void generateCreditNote(
            String templatePath, String outputPath,
            String creditNoteNo, String documentDate, String interest,
            String insured, String reinsurer, String period,
            double sumInsured, double rate, double facPremiumFull,
            double share, double sharePremium, double grossBrokerage, double netPayable)
            throws IOException, InvalidFormatException {

        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            XWPFTable table = doc.getTables().get(0);

            // fill template rows according to the structure you provided earlier
            setCellText(table, 0, 2, creditNoteNo);
            setCellText(table, 1, 2, documentDate);
            // blank row assumed at index 2
            setCellText(table, 3, 2, interest);
            setCellText(table, 4, 2, insured);
            setCellText(table, 5, 2, reinsurer);
            setCellText(table, 6, 2, period);
            setCellText(table, 7, 2, "USD " + format(sumInsured));
            setCellText(table, 8, 2, String.format("%.2f%%", rate));
            setCellText(table, 9, 2, "USD " + format(facPremiumFull));
            setCellText(table, 10, 2, String.format("%.2f%% of 100%%", share));
            setCellText(table, 11, 2, "USD " + format(sharePremium));
            setCellText(table, 12, 2, "USD " + format(grossBrokerage));
            setCellText(table, 13, 2, "USD " + format(netPayable));

            File outFile = new File(outputPath);
            outFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                doc.write(fos);
            }
        }
    }

    private static String format(double val) {
        return String.format("%,.2f", val);
    }

    // --- Auto open folder ---
    private static void openOutputFolder(String outputFolder) {
        try {
            File folder = new File(outputFolder);
            if (folder.exists()) {
                new ProcessBuilder("explorer.exe", folder.getAbsolutePath()).start();
            }
        } catch (IOException e) {
            System.err.println("⚠️ Unable to open output folder automatically.");
        }
    }
}
