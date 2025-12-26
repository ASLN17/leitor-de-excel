import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' hide Border;
import 'package:flutter/services.dart';
import 'package:clipboard/clipboard.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Leitor de Anexo Excel',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: const ExcelToSQLScreen(),
    );
  }
}

class ExcelToSQLScreen extends StatefulWidget {
  const ExcelToSQLScreen({super.key});

  @override
  _ExcelToSQLScreenState createState() => _ExcelToSQLScreenState();
}

class _ExcelToSQLScreenState extends State<ExcelToSQLScreen> {
  String? _selectedFilePath;
  List<String> _sqlScripts = [];
  bool _isLoading = false;

  Future<void> _pickExcelFile() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (result != null) {
        setState(() {
          _selectedFilePath = result.files.single.path;
          _sqlScripts = [];
        });
      }
    } catch (e) {
      _showError('Erro ao selecionar arquivo: $e');
    }
  }

  Future<void> _processExcelFile() async {
    if (_selectedFilePath == null) {
      _showError('Por favor, selecione um arquivo Excel primeiro');
      return;
    }

    setState(() {
      _isLoading = true;
      _sqlScripts = [];
    });

    try {
      // Lê o arquivo Excel
      var bytes = File(_selectedFilePath!).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      // Procura pela planilha "tabela geral"
      Sheet? sheet;
      for (var sheetName in excel.tables.keys) {
        if (sheetName.toLowerCase().contains('tabela geral')) {
          sheet = excel.tables[sheetName];
          break;
        }
      }

      if (sheet == null) {
        _showError('Planilha "tabela geral" não encontrada');
        setState(() => _isLoading = false);
        return;
      }

      List<String> sqlScripts = [];

      // Índices das colunas baseado no seu arquivo
      int colCodigoLC = 0; // Coluna A
      int colDescricao = 1; // Coluna B
      int colNBS = 2; // Coluna C
      int colPsOnerosa = 4; // Coluna E
      int colAdqExterior = 5; // Coluna F
      int colIndop = 6; // Coluna G
      int colCClassTrib = 8; // Coluna I

      // Variáveis para armazenar os últimos valores não-nulos (células mescladas)
      String? ultimoCodigoLC;
      String? ultimaDescricao;
      String? ultimoNBS;
      String? ultimoPsOnerosa;
      String? ultimoAdqExterior;
      String? ultimoIndop;
      String? ultimoCClassTrib;

      // Processa a partir da linha 2 (índice 1) para pular o cabeçalho
      for (int row = 1; row < sheet.maxRows; row++) {
        // Lê os valores da célula atual
        var cellCodigoLC = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colCodigoLC, rowIndex: row),
        );
        var cellDescricao = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colDescricao, rowIndex: row),
        );
        var cellNBS = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colNBS, rowIndex: row),
        );
        var cellPsOnerosa = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colPsOnerosa, rowIndex: row),
        );
        var cellAdqExterior = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: colAdqExterior,
            rowIndex: row,
          ),
        );
        var cellIndop = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colIndop, rowIndex: row),
        );
        var cellCClassTrib = sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: colCClassTrib, rowIndex: row),
        );

        // Atualiza os valores se não forem nulos (lógica de células mescladas)
        if (cellCodigoLC.value != null &&
            cellCodigoLC.value.toString().trim().isNotEmpty) {
          ultimoCodigoLC = _sanitizeString(cellCodigoLC.value.toString());
        }

        if (cellDescricao.value != null &&
            cellDescricao.value.toString().trim().isNotEmpty) {
          ultimaDescricao = _sanitizeString(cellDescricao.value.toString());
        }

        if (cellNBS.value != null &&
            cellNBS.value.toString().trim().isNotEmpty) {
          ultimoNBS = _sanitizeString(cellNBS.value.toString());
        }

        if (cellPsOnerosa.value != null &&
            cellPsOnerosa.value.toString().trim().isNotEmpty) {
          ultimoPsOnerosa = _sanitizeString(cellPsOnerosa.value.toString());
        }

        if (cellAdqExterior.value != null &&
            cellAdqExterior.value.toString().trim().isNotEmpty) {
          ultimoAdqExterior = _sanitizeString(cellAdqExterior.value.toString());
        }

        if (cellIndop.value != null &&
            cellIndop.value.toString().trim().isNotEmpty) {
          ultimoIndop = _sanitizeString(cellIndop.value.toString());
        }

        if (cellCClassTrib.value != null &&
            cellCClassTrib.value.toString().trim().isNotEmpty) {
          ultimoCClassTrib = _sanitizeString(cellCClassTrib.value.toString());
        }

        // Verifica se temos pelo menos o código NBS para gerar o INSERT
        // (pois cada linha com NBS único parece ser um registro)
        if (ultimoNBS != null && ultimoNBS.isNotEmpty) {
          // Usa os últimos valores conhecidos
          String codigoLC = ultimoCodigoLC ?? '';
          String descricao = ultimaDescricao ?? '';
          String nbs = ultimoNBS;

          // Trata valores booleanos
          bool psOnerosa = (ultimoPsOnerosa ?? '').toUpperCase() == 'S';
          bool adqExterior = (ultimoAdqExterior ?? '').toUpperCase() == 'S';

          String indop = ultimoIndop ?? '';
          String cClassTrib = ultimoCClassTrib ?? '';

          // Cria script SQL
          String sql =
              '''
INSERT INTO nfse_indop (codigo_lc, descricao, nbs, ps_onerosa, adq_exterior, indop, cclasstrib) 
VALUES ('$codigoLC', '${descricao.replaceAll("'", "''")}', '$nbs', $psOnerosa, $adqExterior, '$indop', '$cClassTrib');
        '''
                  .trim();

          sqlScripts.add(sql);
        }
      }

      setState(() {
        _sqlScripts = sqlScripts;
        _isLoading = false;
      });

      if (sqlScripts.isEmpty) {
        _showError('Nenhum dado encontrado para processar');
      }
    } catch (e) {
      setState(() => _isLoading = false);
      _showError('Erro ao processar arquivo: $e');
    }
  }

  String _sanitizeString(String input) {
    return input.trim().replaceAll('\n', ' ').replaceAll('\r', '');
  }

  void _showError(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message), backgroundColor: Colors.red),
    );
  }

  void _copyToClipboard() {
    if (_sqlScripts.isEmpty) return;

    String allScripts = _sqlScripts.join('\n\n');
    FlutterClipboard.copy(allScripts).then((_) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('Scripts copiados para a área de transferência!'),
          backgroundColor: Colors.green,
        ),
      );
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Gerador de SQL a partir do Anexo')),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            // Seção de seleção de arquivo
            Card(
              child: Padding(
                padding: const EdgeInsets.all(16.0),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'Selecione o arquivo Excel (Anexo VIII)',
                      style: TextStyle(
                        fontSize: 16,
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                    const SizedBox(height: 10),
                    Text(
                      _selectedFilePath ?? 'Nenhum arquivo selecionado',
                      style: TextStyle(
                        color: _selectedFilePath != null
                            ? Colors.green
                            : Colors.grey,
                      ),
                    ),
                    const SizedBox(height: 15),
                    Row(
                      children: [
                        Expanded(
                          child: ElevatedButton.icon(
                            icon: const Icon(Icons.folder_open),
                            label: const Text('Selecionar Arquivo'),
                            onPressed: _pickExcelFile,
                          ),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: ElevatedButton.icon(
                            icon: _isLoading
                                ? const SizedBox(
                                    width: 16,
                                    height: 16,
                                    child: CircularProgressIndicator(
                                      strokeWidth: 2,
                                      valueColor: AlwaysStoppedAnimation(
                                        Colors.white,
                                      ),
                                    ),
                                  )
                                : const Icon(Icons.play_arrow),
                            label: _isLoading
                                ? const Text('Processando...')
                                : const Text('Gerar SQL'),
                            onPressed: _isLoading ? null : _processExcelFile,
                          ),
                        ),
                      ],
                    ),
                  ],
                ),
              ),
            ),

            const SizedBox(height: 20),

            // Seção de resultados
            Expanded(
              child: _sqlScripts.isEmpty
                  ? Center(
                      child: Text(
                        _isLoading
                            ? 'Processando arquivo...'
                            : 'Selecione um arquivo e clique em "Gerar SQL"',
                        style: TextStyle(fontSize: 18, color: Colors.grey[600]),
                      ),
                    )
                  : Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        Row(
                          mainAxisAlignment: MainAxisAlignment.spaceBetween,
                          children: [
                            Text(
                              'Scripts SQL Gerados (${_sqlScripts.length} registros)',
                              style: const TextStyle(
                                fontSize: 16,
                                fontWeight: FontWeight.bold,
                              ),
                            ),
                            ElevatedButton.icon(
                              icon: const Icon(Icons.content_copy),
                              label: const Text('Copiar Todos'),
                              onPressed: _copyToClipboard,
                            ),
                          ],
                        ),
                        const SizedBox(height: 10),
                        Expanded(
                          child: Container(
                            decoration: BoxDecoration(
                              border: Border.all(color: Colors.grey[300]!),
                              borderRadius: BorderRadius.circular(8),
                            ),
                            child: ListView.builder(
                              itemCount: _sqlScripts.length,
                              itemBuilder: (context, index) {
                                return Container(
                                  padding: const EdgeInsets.all(12),
                                  decoration: BoxDecoration(
                                    border: index > 0
                                        ? Border(
                                            top: BorderSide(
                                              color: Colors.grey[200]!,
                                            ),
                                          )
                                        : null,
                                  ),
                                  child: SelectableText(
                                    _sqlScripts[index],
                                    style: const TextStyle(
                                      fontFamily: 'monospace',
                                      fontSize: 12,
                                    ),
                                  ),
                                );
                              },
                            ),
                          ),
                        ),
                      ],
                    ),
            ),
          ],
        ),
      ),
    );
  }
}
