import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ConcatenateExcelFiles {

    public static void main(String[] args) {
        String inputDirectory = "C:/Users/Luiz Miguel/Downloads/relatorio-automatizado-java/src/main/java/relatorios"; // Substitua pelo caminho real dos seus arquivos
        String outputFile = "C:/Users/Luiz Miguel/Downloads/relatorio-automatizado-java/src/main/java/relatoriosPassados/Relatorio-semanal-29-07a02-08-2024.xlsx";

        try {
            List<Path> excelFiles = listExcelFiles(inputDirectory);
            Collections.sort(excelFiles);
            concatenateExcelFiles(excelFiles, outputFile);
            System.out.println("Arquivos concatenados com sucesso em " + outputFile);
        } catch (IOException e) {
            System.err.println("Erro ao concatenar arquivos: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static List<Path> listExcelFiles(String directory) throws IOException {
        List<Path> excelFiles = new ArrayList<>();
        Files.newDirectoryStream(Paths.get(directory), path -> path.toString().endsWith(".xlsx"))
                .forEach(excelFiles::add);
        return excelFiles;
    }

    public static void concatenateExcelFiles(List<Path> excelFiles, String outputFile) throws IOException {
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Concatenated");

        int rowCount = 0;
        for (Path file : excelFiles) {
            try (FileInputStream fis = new FileInputStream(file.toFile())) {
                Workbook inputWorkbook = new XSSFWorkbook(fis);
                Sheet inputSheet = inputWorkbook.getSheetAt(0);
                for (Row inputRow : inputSheet) {
                    Row outputRow = outputSheet.createRow(rowCount++);
                    copyRow(inputRow, outputRow);
                }
            }
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            outputWorkbook.write(fos);
        }

        outputWorkbook.close();
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex(), sourceCell.getCellType());
            switch (sourceCell.getCellType()) {
                case STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(sourceCell)) {
                        targetCell.setCellValue(sourceCell.getDateCellValue());
                    } else {
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    targetCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                default:
                    break;
            }
        }
    }
}
