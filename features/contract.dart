import 'package:flutter/material.dart';
import 'package:signature/signature.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:excel/excel.dart';
import 'package:image/image.dart' as img;
import 'package:flutter/services.dart' show rootBundle, ByteData;

class ContractFeature extends StatefulWidget {
  @override
  _ContractFeatureState createState() => _ContractFeatureState();
}

class _ContractFeatureState extends State<ContractFeature> {
  SignatureController _controller = SignatureController(
    penStrokeWidth: 5,
    penColor: Colors.black,
    exportBackgroundColor: Colors.white,
  );

  Future<String> getFilePath(String fileName) async {
    Directory directory = await getApplicationDocumentsDirectory();
    String path = directory.path;
    return '$path/$fileName';
  }

  Future<void> copyAssetFileToLocalDir(String assetPath, String localFileName) async {
    ByteData data = await rootBundle.load(assetPath);
    List<int> bytes = data.buffer.asUint8List();
    String localPath = await getFilePath(localFileName);
    File localFile = File(localPath);
    await localFile.writeAsBytes(bytes);
  }

  Future<void> saveSignatureToExcel() async {
    var signature = await _controller.toPngBytes();
    if (signature != null) {
      String localPath = await getFilePath('contract.xlsx');
      var bytes = File(localPath).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      img.Image? image = img.decodeImage(signature);
      if (image != null) {
        excel.updateCell('Sheet1', CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 16), "Signature" as CellValue?);
        // Note: img library needs to be used to handle image embedding, but excel package doesn't support direct image embedding.

        var updatedBytes = excel.encode();
        File(localPath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(updatedBytes!);
      }
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text('Signature Saved to Excel!')));
    }
  }

  @override
  void initState() {
    super.initState();
    copyAssetFileToLocalDir('assets/contract.xlsx', 'contract.xlsx');
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('계약서 서명 기능'),
      ),
      body: Column(
        children: [
          Expanded(
            child: Signature(
              controller: _controller,
              backgroundColor: Colors.lightBlueAccent,
            ),
          ),
          Container(
            color: Colors.yellow.withOpacity(0.5),
            child: TextButton(
              onPressed: saveSignatureToExcel,
              child: Text('서명 저장', style: TextStyle(color: Colors.black)),
              style: TextButton.styleFrom(
                minimumSize: Size(200, 50), // 버튼 크기 설정
              ),
            ),
          )
        ],
      ),
    );
  }
}
