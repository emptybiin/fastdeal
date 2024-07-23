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
  // 각 입력 필드를 제어하기 위한 컨트롤러 생성
  List<TextEditingController> controllers = List.generate(19, (index) => TextEditingController());
  // 추가 입력 필드에 대한 컨트롤러 생성
  List<TextEditingController> extraFieldControllers = List.generate(2, (index) => TextEditingController());
  int currentIndex = 0;

  // 각 필드 이름과 엑셀 셀 위치 정의
  List<String> fieldNames = [
    "고객명", "차량명", "리스사명", "실행회수", "납입횟수",
    "최종납입날짜", "차량매매가", "미회수원금", "보증금", "선납금",
    "잔존가치", "리스료", "일할차세", "일할이자", "승계수수료",
    "판매수수료", "기타비용", "추가 입력 1", "추가 입력 2"
  ];
  List<String> cellPositions = [
    "N1", "N3", "N5", "N7", "N9",
    "N11", // 최종납입날짜
    "N13", "N15", "N17", "N19",
    "N21", "N23", "N25", "N27", "N29",
    "N31", "N33", "N35", "N37"
  ];

  // 파일 경로를 얻기 위한 함수
  Future<String> getFilePath(String fileName) async {
    Directory directory = await getApplicationDocumentsDirectory();
    String path = directory.path;
    return '$path/$fileName';
  }

  // 에셋 파일을 로컬 디렉토리로 복사하는 함수
  Future<void> copyAssetFileToLocalDir(String assetPath, String localFileName) async {
    ByteData data = await rootBundle.load(assetPath);
    List<int> bytes = data.buffer.asUint8List();
    String localPath = await getFilePath(localFileName);
    File localFile = File(localPath);
    await localFile.writeAsBytes(bytes);
  }

  // 입력된 데이터를 엑셀 파일에 저장하는 함수
  Future<void> saveToExcel() async {
    String localPath = await getFilePath('report.xlsx');
    var bytes = File(localPath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var sheet = excel['Sheet1'];

    // 각 컨트롤러의 값을 엑셀 셀에 업데이트
    for (int i = 0; i < controllers.length; i++) {
      String cellValue = controllers[i].text;

      // 최종납입날짜 (N11) 처리는 날짜 형식으로 변환
      if (i == 5 && RegExp(r'^\d{8}$').hasMatch(cellValue)) {
        int year = int.parse(cellValue.substring(0, 4));
        int month = int.parse(cellValue.substring(4, 6));
        int day = int.parse(cellValue.substring(6, 8));
        var dateCell = sheet.cell(CellIndex.indexByString(cellPositions[i]));
        dateCell.value = DateCellValue(year: year, month: month, day: day);
        dateCell.cellStyle = CellStyle(
          numberFormat: NumFormat.standard_14, // 'm/d/yy' 형식
        );
      } else if (i == 17) {
        // 18번째 필드 처리
        if (controllers[i].text.isNotEmpty && extraFieldControllers[0].text.isNotEmpty) {
          sheet.updateCell(CellIndex.indexByString("L35"), extraFieldControllers[0].text as CellValue?);
          sheet.updateCell(CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
        }
      } else if (i == 18) {
        // 19번째 필드 처리
        if (controllers[i].text.isNotEmpty && extraFieldControllers[1].text.isNotEmpty) {
          sheet.updateCell(CellIndex.indexByString("L37"), extraFieldControllers[1].text as CellValue?);
          sheet.updateCell(CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
        }
      } else {
        // 일반 필드 처리
        sheet.updateCell(CellIndex.indexByString(cellPositions[i]), cellValue as CellValue?);
      }
    }

    // 업데이트된 엑셀 파일을 저장
    var updatedBytes = excel.encode();
    String newFilePath = await getFilePath('new_report.xlsx');
    File(newFilePath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(updatedBytes!);

    // 저장 완료 메시지 표시
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text('엑셀 파일이 저장되었습니다: new_report.xlsx')));

    // 파일 공유
    Share.shareFiles([newFilePath], text: '새로운 차량 매매 보고서 엑셀 파일');
  }

  // 초기화 시 에셋 파일을 로컬 디렉토리로 복사
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
              // 현재 필드 이름 표시
              Text(
                '${fieldNames[currentIndex]} (Field ${currentIndex + 1})',
                style: TextStyle(fontSize: 24),
              ),
              SizedBox(height: 20),
              // 추가 입력 필드 처리
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
              // 현재 입력 필드
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
              // 이전, 다음, 완료 버튼
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