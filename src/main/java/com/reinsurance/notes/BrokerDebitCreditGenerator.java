package com.reinsurance.notes;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

public class BrokerDebitCreditGenerator {

    public static void main(String[] args) {

        String basePath = System.getProperty("user.dir");

        String excelFilePath     = basePath + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
        String templatePath      = basePath + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
        String creditTemplatePath= basePath + File.separator + "resources" + File.separator + "CreditNoteTemplate.docx";
        String outputFolder      = basePath + File.separator + "resources" + File.separator + "output" + File.separator;

        File excelFile = new File(excelFilePath);
        if (!excelFile.exists()) {
            excelFilePath      = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
            templatePath       = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
            creditTemplatePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "CreditNoteTemplate.docx";
            outputFolder       = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "output" + File.separator;
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

            Sheet mainSheet   = wb.getSheetAt(0);
            Sheet creditSheet = wb.getSheet("CreditNoteDetails"); // optional

            File outDir = new File(outputFolder);
            if (!outDir.exists()) outDir.mkdirs();

            DateTimeFormatter df = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
            int mainProcessedCol = 21; // unchanged

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

                // --- Inputs (main sheet stays index-based as in your code) ---
                String debitNoteNo = getString(row, 0);
                if (debitNoteNo == null || debitNoteNo.isEmpty()) {
                    debitNoteNo = "DN-" + String.format("%03d", r);
                }

                String docDate = getString(row, 1);
                if (docDate == null || docDate.trim().isEmpty()) {
                    docDate = LocalDate.now().format(df);
                    setString(row, 1, docDate); // persist auto-date
                }

                String interest          = getString(row, 2);
                String insured           = getString(row, 3);
                String defaultReinsured  = getString(row, 4); // shows as "Reinsured" on debit note
                String period            = getString(row, 5);

                double SI             = getDouble(row, 6);
                double cedentRate     = getDouble(row, 7);
                double mainReinsRate  = getDouble(row, 8);
                double share          = getDouble(row, 9);
                double brokerage      = getDouble(row,10);
                double cedingCommPct  = getDouble(row,15);

                if (SI == 0 || cedentRate == 0 || share == 0) {
                    System.out.println("⚠️ Skipping incomplete main row " + r);
                    continue;
                }

                // --- Calculations (unchanged logic) ---
                double grossPremiumCedent   = SI * (cedentRate / 100);
                double sharePremiumCedent   = grossPremiumCedent * (share / 100);

                double effectiveReinsRateMain = (mainReinsRate > 0) ? mainReinsRate : cedentRate;
                double grossPremiumReinsMain  = SI * (effectiveReinsRateMain / 100);
                double sharePremiumReinsMain  = grossPremiumReinsMain * (share / 100);

                double cedingCommissionAmtMain = sharePremiumCedent * (cedingCommPct / 100);
                double grossBrokerageMain      = sharePremiumReinsMain * (brokerage / 100);
                double netBrokerageMain        = grossBrokerageMain / 2.0;

                double netPremiumFromYou = sharePremiumCedent - cedingCommissionAmtMain;                 // debit (cedent)
                double netPremiumToYou   = sharePremiumReinsMain - netBrokerageMain - cedingCommissionAmtMain; // credit (summary)

                // --- Write back (same column indices) ---
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
                setString (row, mainProcessedCol, "Yes");

                // --- Generate Debit Note (unchanged) ---
                String safeFileName = debitNoteNo.replaceAll("[^a-zA-Z0-9-_]", "_");
                generateDebitNote(
                        templatePath,
                        outputFolder + safeFileName + ".docx",
                        debitNoteNo,
                        docDate,
                        interest,
                        insured,
                        defaultReinsured,
                        period,
                        SI,
                        cedentRate,
                        grossPremiumCedent,
                        share,
                        sharePremiumCedent,
                        netPremiumFromYou
                );
                System.out.println("✅ Main Debit Note generated: " + debitNoteNo);

                // --- Credit notes (by header names; robust to column order & new 'Reinsurer Address') ---
                if (creditSheet != null) {
                    Map<String,Integer> hdr = readHeaderMap(creditSheet);
                    for (int cr = 1; cr <= creditSheet.getLastRowNum(); cr++) {
                        Row crow = creditSheet.getRow(cr);
                        if (crow == null) continue;

                        String linkedDebit = getStringByHeader(crow, hdr, "Debit Note No.");
                        if (linkedDebit == null || !linkedDebit.equalsIgnoreCase(debitNoteNo)) continue;

                        String creditProcessed = getStringByHeader(crow, hdr, "Processed");
                        if ("yes".equalsIgnoreCase(creditProcessed) || "processed".equalsIgnoreCase(creditProcessed)) {
                            System.out.println("⏩ Skipping credit row " + cr + " (already processed)");
                            continue;
                        }

                        String creditNoteNo    = getStringByHeader(crow, hdr, "Credit Note No.");
                        String reinsuredName   = getStringByHeader(crow, hdr, "Reinsured");          // inside table
                        String reinsurerName   = getStringByHeader(crow, hdr, "Reinsurer Name");     // To, block
                        String reinsurerAddr   = getStringByHeader(crow, hdr, "Reinsurer Address");  // To, block (optional)

                        double reinsurerShare  = getDoubleByHeader(crow, hdr, "Reinsurer Share (%)");
                        double creditRowRate   = getDoubleByHeader(crow, hdr, "Reinsurance Rate (%)");
                        double creditRowBrok   = getDoubleByHeader(crow, hdr, "Brokerage (%)");
                        double creditRowCedPct = getDoubleByHeader(crow, hdr, "Ceding Commission (%)");

                        if (reinsuredName == null || reinsuredName.trim().isEmpty()) {
                            reinsuredName = defaultReinsured; // fallback
                        }
                        if (reinsurerName == null || reinsurerName.trim().isEmpty()) {
                            reinsurerName = "(Reinsurer)";
                        }

                        double effectiveReinsRate = (creditRowRate  > 0) ? creditRowRate  : ((mainReinsRate > 0) ? mainReinsRate : cedentRate);
                        double effectiveBrokerage = (creditRowBrok  > 0) ? creditRowBrok  : brokerage;
                        double effectiveCedingPct = (creditRowCedPct> 0) ? creditRowCedPct: cedingCommPct;

                        double gpReins    = SI * (effectiveReinsRate / 100);
                        double spReins    = gpReins * (reinsurerShare / 100);
                        double ccAmt      = spReins * (effectiveCedingPct / 100);
                        double gb         = spReins * (effectiveBrokerage / 100);
                        double nb         = gb / 2.0;
                        double netPayable = spReins - nb - ccAmt;

                        // write outputs back if headers exist
                        writeIfPresent(crow, hdr, "Fac Premium 100%",         gpReins);
                        writeIfPresent(crow, hdr, "Share Premium",             spReins);
                        writeIfPresent(crow, hdr, "Ceding Commission (Amount)",ccAmt);
                        writeIfPresent(crow, hdr, "Gross Brokerage",           gb);
                        writeIfPresent(crow, hdr, "Net Premium Payable To You",netPayable);

                        String useCreditNo  = (creditNoteNo != null && !creditNoteNo.trim().isEmpty())
                                ? creditNoteNo
                                : ("CN-" + debitNoteNo + "-" + reinsurerName.replaceAll("[^a-zA-Z0-9]", ""));
                        String safeCredit   = useCreditNo.replaceAll("[^a-zA-Z0-9-_\\.]", "_");

                        try {
                            generateCreditNote(
                                    creditTemplatePath,
                                    outputFolder + safeCredit + ".docx",
                                    useCreditNo,
                                    docDate,
                                    interest,
                                    insured,
                                    reinsuredName,    // table: Reinsured
                                    reinsurerName,    // To, block: name
                                    reinsurerAddr,    // To, block: address (optional)
                                    period,
                                    SI,
                                    effectiveReinsRate,
                                    gpReins,
                                    reinsurerShare,
                                    spReins,
                                    gb,
                                    netPayable
                            );
                            // mark processed if header exists
                            writeIfPresent(crow, hdr, "Processed", "Yes");
                            System.out.println("   ✅ Credit Note generated for " + reinsurerName + " (linked to " + debitNoteNo + ") - CN: " + useCreditNo);
                        } catch (Exception ce) {
                            System.err.println("   ❌ Failed to generate credit note for " + reinsurerName + ": " + ce.getMessage());
                            ce.printStackTrace();
                        }
                    }
                }
            }

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

    // ---------- Helpers: rows/headers ----------

    private static Map<String,Integer> readHeaderMap(Sheet sheet) {
        Map<String,Integer> map = new HashMap<>();
        Row hdr = sheet.getRow(0);
        if (hdr == null) return map;
        for (int i = 0; i < hdr.getLastCellNum(); i++) {
            String key = getString(hdr, i).trim();
            if (!key.isEmpty()) map.put(key.toLowerCase(), i);
        }
        return map;
    }

    private static String getStringByHeader(Row row, Map<String,Integer> hdr, String name) {
        Integer idx = hdr.get(name.toLowerCase());
        return (idx == null) ? "" : getString(row, idx);
    }

    private static double getDoubleByHeader(Row row, Map<String,Integer> hdr, String name) {
        Integer idx = hdr.get(name.toLowerCase());
        return (idx == null) ? 0.0 : getDouble(row, idx);
    }

    private static void writeIfPresent(Row row, Map<String,Integer> hdr, String name, double val) {
        Integer idx = hdr.get(name.toLowerCase());
        if (idx != null) setNumeric(row, idx, val);
    }

    private static void writeIfPresent(Row row, Map<String,Integer> hdr, String name, String val) {
        Integer idx = hdr.get(name.toLowerCase());
        if (idx != null) setString(row, idx, val);
    }

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

    // ---------- Excel low-level helpers ----------

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
            try { return Double.parseDouble(c.getStringCellValue().replace(",", "").trim()); }
            catch (Exception e) { return 0.0; }
        }
        return 0.0;
    }

    private static void setNumeric(Row row, int idx, double val) {
        if (idx > 200) return;
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.NUMERIC);
        c.setCellValue(val);
    }

    private static void setString(Row row, int idx, String val) {
        if (idx > 200) return;
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.STRING);
        c.setCellValue(val);
    }

    // ---------- Word helpers ----------

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

    // Replace the “To,” block with Name + Address
    private static void replaceToBlock(XWPFDocument doc, String reinsurerName, String reinsurerAddress) {
        for (XWPFParagraph para : doc.getParagraphs()) {
            String text = para.getText();
            if (text != null && text.trim().toLowerCase().startsWith("to")) {
                for (int i = para.getRuns().size() - 1; i >= 0; i--) {
                    para.removeRun(i);
                }
                XWPFRun run = para.createRun();
                run.setText("To,");
                run.addBreak();
                if (reinsurerName != null && !reinsurerName.isEmpty()) {
                    run.setText(reinsurerName);
                    if (reinsurerAddress != null && !reinsurerAddress.trim().isEmpty()) {
                        run.addBreak();
                        // support multi-line address separated by \n
                        String[] lines = reinsurerAddress.split("\\r?\\n");
                        for (int i = 0; i < lines.length; i++) {
                            if (i > 0) run.addBreak();
                            run.setText(lines[i]);
                        }
                    }
                }
                return;
            }
        }
    }

    private static void generateCreditNote(
            String templatePath, String outputPath,
            String creditNoteNo, String documentDate, String interest,
            String insured, String reinsuredName, String reinsurerName, String reinsurerAddress,
            String period,
            double sumInsured, double rate, double facPremiumFull,
            double share, double sharePremium, double grossBrokerage, double netPayable)
            throws IOException, InvalidFormatException {

        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // Update the “To,” block (name + optional address)
            replaceToBlock(doc, reinsurerName, reinsurerAddress);

            XWPFTable table = doc.getTables().get(0);

            setCellText(table, 0, 2, creditNoteNo);
            setCellText(table, 1, 2, documentDate);
            setCellText(table, 3, 2, interest);
            setCellText(table, 4, 2, insured);
            setCellText(table, 5, 2, reinsuredName); // Reinsured inside table
            setCellText(table, 6, 2, period);
            setCellText(table, 7, 2, "USD " + format(sumInsured));
            setCellText(table, 8, 2, String.format("%.2f%%", rate));
            setCellText(table, 9, 2, "USD " + format(facPremiumFull));
            setCellText(table,10, 2, String.format("%.2f%% of 100%%", share));
            setCellText(table,11, 2, "USD " + format(sharePremium));
            setCellText(table,12, 2, "USD " + format(grossBrokerage));
            setCellText(table,13, 2, "USD " + format(netPayable));

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
