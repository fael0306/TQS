import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class TQSExcelProcessor extends JFrame {

    private JTextArea outputArea;

    public TQSExcelProcessor() {
        setTitle("Processador de Arquivo TQS");
        setSize(400, 300);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JButton selectFileButton = new JButton("Selecionar Arquivo TQS");
        selectFileButton.addActionListener(e -> selecionarArquivo());

        outputArea = new JTextArea();
        outputArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(outputArea);

        add(selectFileButton, BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);
    }

    private void selecionarArquivo() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Selecione o arquivo TQS");

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            outputArea.setText("Arquivo selecionado: " + selectedFile.getAbsolutePath());
            processarArquivo(selectedFile);
        } else {
            outputArea.setText("Nenhum arquivo foi selecionado.");
        }
    }

    private List<String> lerArquivo(File file) {
        List<String> linhas = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file), "ISO-8859-1"))) {
            String linha;
            while ((linha = reader.readLine()) != null) {
                linhas.add(linha);
            }
        } catch (FileNotFoundException e) {
            outputArea.setText("Erro: O arquivo não foi encontrado.");
        } catch (IOException e) {
            outputArea.setText("Erro ao ler o arquivo.");
        }
        return linhas;
    }

    private void processarArquivo(File file) {
        List<String> linhas = lerArquivo(file);
        if (linhas.isEmpty()) {
            outputArea.append("\nArquivo vazio ou não encontrado.");
            return;
        }
    
        boolean quantitativosEncontrados = false;
        List<String> linhasEncontradas = new ArrayList<>();
    
        for (int i = 0; i < linhas.size(); i++) {
            String linha = linhas.get(i);
            if (linha.contains("Quantitativos")) {
                quantitativosEncontrados = true;
                for (int j = i + 1; j < linhas.size(); j++) {
                    if (linhas.get(j).contains("Legenda")) {
                        break;
                    }
                    linhasEncontradas.add(linhas.get(j).trim());
                }
                break;
            }
        }
    
        if (quantitativosEncontrados) {
            List<List<String>> palavras = new ArrayList<>();
            for (String linha : linhasEncontradas) {
                List<String> palavrasLinha = List.of(linha.split("\\s+"));
                palavras.add(palavrasLinha);
            }
            // Aqui você precisa passar 'file' e 'palavras' para o método salvarComoXLS
            salvarComoXLS(palavras, file); // Chamada corrigida
        } else {
            outputArea.append("\nA palavra 'Quantitativos' não foi encontrada.");
        }
    }    

    private void salvarComoXLS(List<List<String>> palavras, File selectedFile) {
        // Extrai o nome do arquivo original sem a extensão
        String nomeArquivo = selectedFile.getName();
        // Define o novo nome do arquivo com a extensão .xls
        String novoNomeArquivo = nomeArquivo.substring(0, nomeArquivo.lastIndexOf('.')) + "_extraidas.xls";
        File file = new File(novoNomeArquivo);
    
        try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), "UTF-8"))) {
            writer.write("<?xml version=\"1.0\"?>\n");
            writer.write("<?mso-application progid=\"Excel.Sheet\"?>\n");
            writer.write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
            writer.write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
            writer.write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
            writer.write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
            writer.write(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");
            writer.write("<Worksheet ss:Name=\"Sheet1\">\n");
            writer.write("<Table>\n");
    
            for (List<String> linha : palavras) {
                writer.write("  <Row>\n");
                for (String palavra : linha) {
                    // Substitui "." por "," apenas para números
                    palavra = palavra.replace(".", ",");
                    writer.write("    <Cell><Data ss:Type=\"String\">" + palavra + "</Data></Cell>\n");
                }
                writer.write("  </Row>\n");
            }
    
            writer.write("</Table>\n");
            writer.write("</Worksheet>\n");
            writer.write("</Workbook>\n");
    
            outputArea.append("\nDados exportados para '" + file.getAbsolutePath() + "'");
        } catch (IOException e) {
            outputArea.append("\nErro ao salvar o arquivo XLS.");
        }
    }
      
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            TQSExcelProcessor frame = new TQSExcelProcessor();
            frame.setVisible(true);
        });
    }
}