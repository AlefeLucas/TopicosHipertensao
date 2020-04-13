import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

enum Tipo {
    NUMERICO,
    BINARIO,
    NOMINAL,
    OUTRO
}

class Resposta {
    String constante;
    String descricao;
    int count = 0;

    @Override
    public String toString() {
        return "Resposta{" +
                "constante='" + constante + '\'' +
                ", descricao='" + descricao + '\'' +
                '}';
    }
}

class Atributo {
    String id;
    int posicaoInicial = -1;
    int tamanho = -1;
    Tipo tipo;
    int faixaMinima = -1;
    int faixaMaxima = -1;
    String descricao;
    HashMap<String, Resposta> respostas;
    HashMap<String, Resposta> respostasNum;

    public Atributo() {
        respostas = new HashMap<>();
        respostasNum = new HashMap<>();
    }

    @Override
    public String toString() {
        return "Atributo{" +
                "id='" + id + '\'' +
                ", posicaoInicial=" + posicaoInicial +
                ", tamanho=" + tamanho +
                ", tipo=" + tipo +
                ((tipo == Tipo.NUMERICO && faixaMinima != -1) ?
                        ", faixaMinima=" + faixaMinima +
                                ", faixaMaxima=" + faixaMaxima : "") +
                ", descricao='" + descricao + '\'' +
                ", respostas='" + respostas + '\'' +
                '}';
    }
}

class MyInt {
    int number = 0;

    public MyInt(int number) {
        this.number = number;
    }

    public int postIncrement() {
        return number++;
    }
}

public class Main {

    public static final String SAIDA_PESSOAS = "dadosPessoas.xlsx";
    public static final String SAIDA_DOMICILIO = "dadosDomicilio.xlsx";

    public static final String DADOS_PESSOAS = "PESPNS2013.txt";
    public static final String DADOS_DOMICILIO = "DOMPNS2013.txt";

    public static final String DICIONARIO_PESSOAS = "Dicionario_de_variaveis_de_pessoas_PNS_2013.xlsx";
    public static final String DICIONARIO_DOCIMILIO = "Dicionario_de_variaveis_de_domicilios_PNS_2013.xlsx";

    public static final String SHEET_NAME = "PNS 2013";

    public static final String[] ATRIBUTOS_SELECIONADOS_DOMICILIO = {
    };

    public static final String[] ATRIBUTOS_NUMERICOS_SEM_FAIXA = {
            "C008", "P00101", "P00301", "P00401", "P006", "P007", "P009", "P011", "P015", "P016", "P018",
            "P020", "P023", "P025", "P026", "P028", "P029", "P031", "P035", "P03701", "P03702", "P03901", "P03902", "P03903", "P04101", "P04102", "P042", "P04301", "P04302",
            "P04403", "P053", "P05402", "P05405",
    };

    public static final String[] ATRIBUTOS_SELECIONADOS_PESSOAS = {
            "V0001", "C006", "C008", "C009", "D001", "E019", "F001", "F00102", "F007", "F00702",
            "F008", "F00802", "VDF001", "VDF00102", "G002", "G00201", "G003", "G006", "G007", "G00701", "G008",
            "G009", "G021", "G022", "G02201", "G023", "G024", "G02501", "G02502", "G02503", "I001", "I005",
            "J001", "J002", "J004", "J005", "J006", "J007", "J008", "J010", "J011", "J012", "J014", "J015",
            "J016", "J027", "J037", "J038", "J039", "J04001", "J04002", "J053", "J054", "J058", "K045", "K046",
            "K047", "K052", "K053", "M005", "M006", "M007", "M008", "M010", "M01106", "M014", "M016", "N004",
            "N005", "N015", "N016", "O022", "O024", "P00101", "P00301", "P00401", "P005", "P006", "P007", "P009",
            "P011", "P012", "P015", "P016", "P017", "P018", "P019", "P020", "P021", "P022", "P023", "P024",
            "P025", "P026", "P02601", "P028", "P029", "P031", "P032", "P034", "P035", "P036", "P03701", "P03702",
            "P038", "P039", "P03901", "P03902", "P03903", "P040", "P04101", "P04102", "P042", "P04301", "P04302",
            "P04403", "P045", "P046", "P047", "P048", "P050", "P052", "P053", "P05401", "P05402", "P05405",
            "P05408", "P05411", "P05414", "P05417", "P05421", "P05802", "P05901", "P05902", "P05903", "P05904",
            "P060", "P061", "P062", "P066", "Q001", "Q002", "Q003", "Q004", "Q005", "Q006", "Q007", "Q008",
            "Q009", "Q010", "Q011", "Q017", "Q01801", "Q01802", "Q01803", "Q01804", "Q01805", "Q01806", "Q01807",
            "Q01808", "Q01901", "Q01902", "Q01903", "Q01904", "Q01905", "Q026", "Q027", "Q028", "Q030", "Q031",
            "Q032", "Q03401", "Q03402", "Q039", "Q053", "Q054", "Q05504", "Q05505", "Q060", "Q061", "Q064",
            "Q06501", "Q06502", "Q06503", "Q066", "Q067", "Q068", "Q069", "Q070", "Q071", "Q07201", "Q07202",
            "Q07203", "Q07204", "Q07205", "Q088", "Q089", "Q09001", "Q092", "Q093", "Q094", "Q09601", "Q09602",
            "Q097", "Q109", "Q110", "Q11001", "Q11002", "Q11004", "Q111", "Q112", "Q116", "Q11601", "Q11602",
            "Q11603", "Q117", "Q11801", "Q120", "Q121", "Q122", "Q124", "Q125", "Q12601", "Q12602", "Q127",
            "Q128", "Q130", "Q131", "Q132", "Q133", "Q134", "Q136", "Q137", "R028", "R029", "S001", "S004",
            "S01002", "S01003", "S01004", "S01103", "S01401", "S01402", "S015", "S016", "S017", "S021", "W00103",
            "W00203", "W00303", "W00407", "W00408", "VDD004"
    };


    public static void main(String[] args) throws IOException {
        File dicionario = new File(DICIONARIO_PESSOAS);

        List<Atributo> atributos = getAtributos(dicionario);
        atributos = atributos.stream().filter(atributo -> Arrays.asList(ATRIBUTOS_SELECIONADOS_PESSOAS).contains(atributo.id)).collect(Collectors.toList());

        File saida = new File("dataPessoas.xlsx");

        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        createHeader(atributos, sheet);

        fillData(atributos, sheet);

        Workbook wMeta = new SXSSFWorkbook();
        Sheet meta = wMeta.createSheet("ESTATISTICAS");


        processMetaData(wMeta, atributos, meta);
        File atributosSaida = new File("atributosPessoas.xlsx");
        save(atributosSaida, wMeta);
        //save(saida, workbook);

    }

    private static void processMetaData(Workbook wMeta, List<Atributo> atributos, Sheet meta) {
        MyInt lastRowWritten = new MyInt(0);

        for (Atributo atributo : atributos) {
            meta.createRow(lastRowWritten.postIncrement());
            Row header1 = meta.createRow(lastRowWritten.postIncrement());
            header1.createCell(0).setCellValue(atributo.id);
            header1.createCell(1).setCellValue(atributo.tipo.name());
            setBorderCell(wMeta, header1.getCell(0));
            setBorderCell(wMeta, header1.getCell(1));

            Row header2 = meta.createRow(lastRowWritten.postIncrement());
            header2.createCell(0).setCellValue(atributo.descricao);
            setBorderCell(wMeta, header2.getCell(0));
            setBorderCell(wMeta, header2.createCell(1));
            if (atributo.tipo != Tipo.NUMERICO && atributo.tipo != Tipo.OUTRO) {
                setBorderCell(wMeta, header1.createCell(2));
                header2.getCell(1).setCellValue("Total");
                header2.createCell(2).setCellValue("Percentual");
                setBorderCell(wMeta, header2.getCell(1));
                setBorderCell(wMeta, header2.getCell(2));
                Optional<Integer> soma = atributo.respostas.values().stream().map(resposta -> resposta.count).reduce(Integer::sum);
                atributo.respostas.forEach((s, resposta) -> {
                            Row currentRow = meta.createRow(lastRowWritten.postIncrement());
                            currentRow.createCell(0).setCellValue(resposta.descricao);
                            currentRow.createCell(1).setCellValue(resposta.count);
                            currentRow.createCell(2).setCellValue(resposta.count / (double) soma.orElseThrow());
                            setBorderCell(wMeta, currentRow.getCell(0));
                            setBorderCell(wMeta, currentRow.getCell(1));
                            setBorderCell(wMeta, currentRow.getCell(2));
                        }
                );
            }

            if (atributo.tipo.equals(Tipo.NUMERICO)) {
                DescriptiveStatistics stats = new DescriptiveStatistics();
                atributo.respostasNum.values().stream().filter(resposta -> !resposta.descricao.isBlank()).forEach(resposta -> {
                    double v = Double.parseDouble(resposta.descricao);
                    for (int i = 0; i < resposta.count; i++) {
                        stats.addValue(v);
                    }
                });

                double mean = stats.getMean();
                double standardDeviation = stats.getStandardDeviation();
                double min = stats.getMin();
                double max = stats.getMax();


                Row row = meta.createRow(lastRowWritten.postIncrement());
                row.createCell(0).setCellValue("Média");
                row.createCell(1).setCellValue(mean);
                setBorderCell(wMeta, row.getCell(0));
                setBorderCell(wMeta, row.getCell(1));

                row = meta.createRow(lastRowWritten.postIncrement());
                row.createCell(0).setCellValue("Desvio Padrão");
                row.createCell(1).setCellValue(standardDeviation);
                setBorderCell(wMeta, row.getCell(0));
                setBorderCell(wMeta, row.getCell(1));

                row = meta.createRow(lastRowWritten.postIncrement());
                row.createCell(0).setCellValue("Mínimo");
                row.createCell(1).setCellValue(min);
                setBorderCell(wMeta, row.getCell(0));
                setBorderCell(wMeta, row.getCell(1));

                row = meta.createRow(lastRowWritten.postIncrement());
                row.createCell(0).setCellValue("Máximo");
                row.createCell(1).setCellValue(max);
                setBorderCell(wMeta, row.getCell(0));
                setBorderCell(wMeta, row.getCell(1));

                row = meta.createRow(lastRowWritten.postIncrement());
                if (atributo.faixaMinima != -1) {
                    row.createCell(0).setCellValue("Faixa");
                    row.createCell(1).setCellValue(atributo.faixaMinima + " a " + atributo.faixaMaxima);
                    setBorderCell(wMeta, row.getCell(0));
                    setBorderCell(wMeta, row.getCell(1));
                }
            }

        }
    }

    private static void setBorderCell(Workbook workbook, Cell cell) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);
    }

    private static void save(File saida, Workbook workbook) throws IOException {
        if (!saida.exists()) {
            saida.createNewFile();
        }

        FileOutputStream outputStream = new FileOutputStream(saida);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    private static void createHeader(List<Atributo> atributos, Sheet sheet) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < atributos.size(); i++) {
            Atributo atributo = atributos.get(i);
            Cell cell = row.createCell(i);
            cell.setCellValue(atributo.id);
        }
    }

    private static void fillData(List<Atributo> atributos, Sheet sheet) throws FileNotFoundException {
        File dados = new File(DADOS_PESSOAS);
        Scanner scanner = new Scanner(dados);
        int linhaCount = 1;
        while (scanner.hasNext()) {
            if (linhaCount % 100 == 0) {
                System.out.println("linha: " + linhaCount);
            }

            linhaCount++;
            String linha = scanner.nextLine();
            Atributo UF = atributos.stream().filter(atributo -> atributo.id.equals("V0001")).collect(Collectors.toList()).get(0);

            String str = linha.substring(UF.posicaoInicial - 1, UF.tamanho + UF.posicaoInicial - 1).trim();
            String str2 = str.equals(".") ? "" : str;
            List<Resposta> rs = UF.respostas.values().stream()
                    .filter(r -> r.constante.equals(str2))
                    .collect(Collectors.toList());
            Resposta resposta = rs.get(0);
            if (Integer.parseInt(resposta.constante) < 31 || Integer.parseInt(resposta.constante) > 43) {
                continue;
            }

            Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);


            for (Atributo atributo : atributos) {
                str = linha.substring(atributo.posicaoInicial - 1, atributo.tamanho + atributo.posicaoInicial - 1).trim();
                final String valor = str.equals(".") ? "" : str;
                int newPosition = newRow.getLastCellNum() == -1 ? 0 : newRow.getLastCellNum();


                switch (atributo.tipo) {
                    case NOMINAL:
                    case BINARIO:
                        rs = atributo.respostas.values().stream()
                                .filter(r -> r.constante.equals(valor))
                                .collect(Collectors.toList());
                        resposta = rs.get(0);
                        resposta.count++;
                        newRow.createCell(newPosition).setCellValue(resposta.descricao);
                        break;
                    case NUMERICO:
                        if (valor.isBlank()) {
                            newRow.createCell(newPosition);
                        } else {
                            newRow.createCell(newPosition).setCellValue(Double.parseDouble(valor));
                        }

                        countRespostasNaoNominais(atributo, valor);


                        break;
                    case OUTRO:
                        if (valor.isBlank()) {
                            newRow.createCell(newPosition);
                        } else {
                            newRow.createCell(newPosition).setCellValue(valor);
                        }

                        countRespostasNaoNominais(atributo, valor);
                        break;
                }

            }
        }
    }

    private static void countRespostasNaoNominais(Atributo atributo, String valor) {
        if (atributo.respostasNum.containsKey(valor.trim())) {
            atributo.respostasNum.get(valor.trim()).count++;
        } else {
            Resposta r = new Resposta();
            r.count = 1;
            r.constante = valor.trim();
            r.descricao = r.constante;
            atributo.respostasNum.put(valor.trim(), r);
        }
    }

    private static List<Atributo> getAtributos(File file) throws IOException {
        FileInputStream inputStream = new FileInputStream(file);
        Sheet sheet = new XSSFWorkbook(inputStream).getSheet(SHEET_NAME);

        List<String> separadores = sheet.getMergedRegions().parallelStream()
                .map(cellAddresses -> sheet.getRow(cellAddresses.getFirstRow()).getCell(0).toString())
                .collect(Collectors.toList());

        List<String> ids = new ArrayList<>(sheet.getLastRowNum() + 1);
        List<Atributo> atributos = new ArrayList<>(sheet.getLastRowNum() + 1);
        for (int i = 3; i <= sheet.getLastRowNum() && sheet.getRow(i).getCell(0) != null; i++) {
            Row row = sheet.getRow(i);
            if (row.getLastCellNum() > 1) {
                Cell firstCell = row.getCell(0);
                if (!firstCell.toString().isBlank() && !separadores.contains(firstCell.toString())) {
                    ids.add(firstCell.getStringCellValue());
                    Atributo atributo = new Atributo();
                    atributo.id = firstCell.getStringCellValue();

                    atributo.posicaoInicial = (int) Double.parseDouble(row.getCell(1).toString());
                    atributo.tamanho = (int) Double.parseDouble(row.getCell(2).toString());
                    atributo.descricao = row.getCell(4).toString();

                    int j;
                    for (j = i + 1; j <= sheet.getLastRowNum() &&
                            sheet.getRow(j).getCell(0).toString().isBlank() &&
                            !sheet.getRow(j).getCell(4).toString().isBlank(); j++)
                        ;
                    int qtdLinhasBrancasAbaixo = j - (i + 1);

                    if (row.getCell(3).toString().contains("dígito")) {
                        atributo.tipo = Tipo.NUMERICO;
                    }

                    for (int k = i + 1; k <= i + qtdLinhasBrancasAbaixo; k++) {
                        Cell cell = sheet.getRow(k).getCell(3);

                        String constante = cell.getCellType() == CellType.NUMERIC ? Integer.toString((int) cell.getNumericCellValue()) : cell.getStringCellValue();
                        Resposta resposta = new Resposta();
                        resposta.constante = constante;
                        resposta.descricao = sheet.getRow(k).getCell(4).toString();
                        atributo.respostas.put(constante, resposta);

                        if (constante.contains(" a ")) {
                            atributo.tipo = Tipo.NUMERICO;
                            String[] ss = constante.split(" a ");
                            atributo.faixaMinima = Integer.parseInt(ss[0]);
                            atributo.faixaMaxima = Integer.parseInt(ss[1]);

                            if (k == i + 2) {
                                String anterior = sheet.getRow(k - 1).getCell(3).toString();

                                if (!anterior.isBlank()) {
                                    atributo.faixaMinima = Integer.parseInt(anterior);
                                }
                            }
                            break;
                        }
                    }

                    if (atributo.tipo == null) {
                        if (atributo.respostas.size() == 2) {
                            atributo.tipo = Tipo.BINARIO;
                        } else if (atributo.respostas.size() == 3 && atributo.respostas.values().stream()
                                .map(resposta -> resposta.descricao).collect(Collectors.toList()).contains("Não aplicável")) {
                            atributo.tipo = Tipo.BINARIO;
                        } else if (atributo.respostas.size() >= 3) {
                            atributo.tipo = Tipo.NOMINAL;
                        } else {
                            atributo.tipo = Tipo.NUMERICO;
                        }
                    }

                    atributos.add(atributo);
                }
            }
        }
        return atributos;
    }
}
