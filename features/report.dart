import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:flutter/services.dart' show rootBundle, ByteData;
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';

class ReportFeature extends StatefulWidget {
  @override
  _ReportFeatureState createState() => _ReportFeatureState();
}

class _ReportFeatureState extends State<ReportFeature> {
  List<TextEditingController> controllers = List.generate(19, (index) => TextEditingController());
  List<TextEditingController> extraFieldControllers = List.generate(2, (index) => TextEditingController());
  int currentIndex = 0;
  List<String> fieldNames = [
    "고객명", "차량명", "리스사명", "실행회수", "납입횟수",
    "최종납입날짜", "차량매매가", "미회수원금", "보증금", "선납금",
    "잔존가치", "리스료", "일할차세", "일할이자", "승계수수료",
    "판매수수료", "기타비용", "추가 입력 1", "추가 입력 2"
  ];
  List<String> cellPositions = [
    "N1", "N3", "N5", "N7", "N9",
    "N11", "N13", "N15", "N17", "N19",
    "N21", "N23", "N25", "N27", "N29",
    "N31", "N33", "N35", "N37"
  ];

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

  Future<void> saveToExcel() async {
    String localPath = await getFilePath('report.xlsx');
    var bytes = File(localPath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    for (int i = 0; i < controllers.length; i++) {
      String cellValue = controllers[i].text;
      if (i == 5) { // 최종납입날짜 처리
        if (RegExp(r'^\d{8}$').hasMatch(cellValue)) {
          cellValue = '${cellValue.substring(0, 4)}.${cellValue.substring(4, 6)}.${cellValue.substring(6, 8)}';
        }
      }

      if (i == 17) {
        // 18번째 필드
        if (controllers[i].text.isNotEmpty && extraFieldControllers[0].text.isNotEmpty) {
          excel.updateCell('Sheet1', CellIndex.indexByString("L35"), extraFieldControllers[0].text as CellValue?);
          excel.updateCell('Sheet1', CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
        }
      } else if (i == 18) {
        // 19번째 필드
        if (controllers[i].text.isNotEmpty && extraFieldControllers[1].text.isNotEmpty) {
          excel.updateCell('Sheet1', CellIndex.indexByString("L37"), extraFieldControllers[1].text as CellValue?);
          excel.updateCell('Sheet1', CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
        }
      } else {
        excel.updateCell('Sheet1', CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
      }
    }

    var updatedBytes = excel.encode();
    String newFilePath = await getFilePath('new_report.xlsx');
    File(newFilePath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(updatedBytes!);

    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text('엑셀 파일이 저장되었습니다: new_report.xlsx')));

    Share.shareFiles([newFilePath], text: '새로운 차량 매매 보고서 엑셀 파일');
  }

  @override
  void initState() {
    super.initState();
    copyAssetFileToLocalDir('assets/report.xlsx', 'report.xlsx');
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('차량 매매 보고서'),
      ),
      body: Center(
        child: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Text(
                '${fieldNames[currentIndex]} (Field ${currentIndex + 1})',
                style: TextStyle(fontSize: 24),
              ),
              SizedBox(height: 20),
              if (currentIndex >= 17 && currentIndex <= 18)
                Column(
                  children: [
                    TextField(
                      controller: extraFieldControllers[currentIndex - 17],
                      style: TextStyle(fontSize: 24),
                      decoration: InputDecoration(
                        hintText: '필드 이름을 입력하세요',
                        border: OutlineInputBorder(),
                      ),
                    ),
                    SizedBox(height: 20),
                  ],
                ),
              Container(
                padding: EdgeInsets.all(16.0),
                decoration: BoxDecoration(
                  color: Colors.grey[800],
                  borderRadius: BorderRadius.circular(12.0),
                ),
                child: TextField(
                  controller: controllers[currentIndex],
                  style: TextStyle(fontSize: 24),
                  decoration: InputDecoration(
                    border: InputBorder.none,
                  ),
                ),
              ),
              SizedBox(height: 20),
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  if (currentIndex > 0)
                    Expanded(
                      child: ElevatedButton(
                        onPressed: () {
                          setState(() {
                            currentIndex--;
                          });
                        },
                        child: Text('이전'),
                      ),
                    ),
                  SizedBox(width: 10),
                  if (currentIndex < fieldNames.length - 1)
                    Expanded(
                      child: ElevatedButton(
                        onPressed: () {
                          setState(() {
                            currentIndex++;
                          });
                        },
                        child: Text('다음'),
                      ),
                    ),
                  SizedBox(width: 10),
                  if (currentIndex == fieldNames.length - 1)
                    Expanded(
                      child: ElevatedButton(
                        onPressed: () {
                          saveToExcel();
                        },
                        child: Text('완료'),
                      ),
                    ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }
}