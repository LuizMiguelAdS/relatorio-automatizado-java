import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class Main {

    public static void main(String[] args) {
        LocalDate dataAtual = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        String nomeArquivoTexto = "C:/Users/Luiz Miguel/Downloads/relatorio-automatizado-java/src/main/java/Relatorio.txt";
        String nomeArquivoExcel = "C:/Users/Luiz Miguel/Downloads/relatorio-automatizado-java/src/main/java/relatorios/Relatorio-" + dataAtual.format(formatter) + ".xlsx";

        try {
            // Leitura do arquivo de texto
            BufferedReader reader = new BufferedReader(new FileReader(nomeArquivoTexto));
            StringBuilder conteudo = new StringBuilder();
            String linha;

            while ((linha = reader.readLine()) != null) {
                conteudo.append(linha).append("\n");
            }

            reader.close();

            // Chama o método para escrever no arquivo Excel
            ExcelWriter.writeExcel(nomeArquivoExcel, conteudo.toString());

            System.out.println("Dados foram registrados no arquivo Excel com sucesso: " + nomeArquivoExcel);

        } catch (IOException e) {
            System.err.println("Erro ao ler o arquivo de texto: " + e.getMessage());
        }
    }
}
