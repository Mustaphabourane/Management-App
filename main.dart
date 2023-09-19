import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';

void main() async {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(
          title: Text('Management App'),
        ),
        body: MyHomePage(),
      ),
    );
  }
}

class MyHomePage extends StatefulWidget {
  @override
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  final List<DataRow> dataRows = [];
  final TextEditingController _dateController = TextEditingController();
  final TextEditingController _vauscherController = TextEditingController();
  final TextEditingController _paxController = TextEditingController();
  final TextEditingController _libelleController = TextEditingController();
  final TextEditingController _amountController = TextEditingController();

  void addData() {
    String date = _dateController.text;
    String vauscher = _vauscherController.text;
    String pax = _paxController.text;
    String libelle = _libelleController.text;
    String amount = _amountController.text;

    DataRow newRow = DataRow(cells: [
      DataCell(Text(date)),
      DataCell(Text(vauscher)),
      DataCell(Text(pax)),
      DataCell(Text(libelle)),
      DataCell(Text(amount)),
    ]);

    setState(() {
      dataRows.add(newRow);
    });

    _dateController.clear();
    _vauscherController.clear();
    _paxController.clear();
    _libelleController.clear();
    _amountController.clear();

    if (dataRows.length >= 2) {
      generateExcel(dataRows);
    }
  }

  Future<void> generateExcel(List<DataRow> dataRows) async {
    // Create a new Excel workbook
    var excel = Excel.createExcel();

    // Add a sheet to the workbook
    Sheet sheetObject = excel['Sheet1'];

    // Add data to the sheet
    sheetObject.appendRow(['Date', 'Vauscher', 'Pax', 'Libelle', 'Amount']);
    for (DataRow row in dataRows) {
      List<String> rowData = [];
      for (DataCell cell in row.cells) {
        rowData.add(cell.child.toString());
      }
      sheetObject.appendRow(rowData);
    }

    // Save the workbook to a file
    
    final directory = await getApplicationDocumentsDirectory();
    final path = '${directory.path}/facture.xlsx';
    

    print('Excel file generated at path: $path');
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: SingleChildScrollView(
        child: Column(
          children: [
            DataTable(
              columns: const [
                DataColumn(label: Text('Date')),
                DataColumn(label: Text('Vauscher')),
                DataColumn(label: Text('Pax')),
                DataColumn(label: Text('Libelle')),
                DataColumn(label: Text('Amount')),
              ],
              rows: dataRows,
            ),
            Padding(
              padding: EdgeInsets.all(10),
              child: Column(
                children: [
                  TextField(
                    controller: _dateController,
                    decoration: const InputDecoration(labelText: 'Date'),
                  ),
                  TextField(
                    controller: _vauscherController,
                    decoration: const InputDecoration(labelText: 'Vauscher'),
                  ),
                  TextField(
                    controller: _paxController,
                    decoration: const InputDecoration(labelText: 'Pax'),
                  ),
                  TextField(
                    controller: _libelleController,
                    decoration: const InputDecoration(labelText: 'Libelle'),
                  ),
                  TextField(
                    controller: _amountController,
                    decoration: const InputDecoration(labelText: 'Amount'),
                  ),
                  ElevatedButton(
                    onPressed: addData,
                    child: const Text('Add Data'),
                  ),
                  ElevatedButton(
                    onPressed: () {
                      generateExcel(dataRows);
                    },
                    child: const Text('Generate Excel'),
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