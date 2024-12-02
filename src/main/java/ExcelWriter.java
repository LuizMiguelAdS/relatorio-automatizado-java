import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void writeExcel(String nomeArquivoExcel, String conteudo) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Relatório");

            // Quebra o conteúdo em registros separados
            String[] registros = conteudo.split("\n\n");

            int rowNum = 0;
            for (String registro : registros) {
                // Processa cada registro
                String[] linhas = registro.split("\n");

                Row row = sheet.createRow(rowNum++);

                int colNum = 0;
                for (String linha : linhas) {
                    String[] partes = linha.split(": ", 2); // Divide apenas na primeira ocorrência de ": "
                    String chave = partes[0];
                    String valor = partes.length > 1 ? partes[1] : "";

                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(valor.trim());
                }
            }

            // Escreve o conteúdo no arquivo Excel
            try (FileOutputStream outputStream = new FileOutputStream(nomeArquivoExcel)) {
                workbook.write(outputStream);
            }

        } catch (IOException e) {
            System.err.println("Erro ao escrever no arquivo Excel: " + e.getMessage());
        }
    }
}
