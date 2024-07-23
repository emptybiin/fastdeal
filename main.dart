import 'package:flutter/material.dart';
import 'features/report.dart';
import 'features/contract.dart';
import 'splash.dart'; // splash.dart 파일 경로 추가

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Flutter Demo',
      theme: ThemeData.dark().copyWith(
        primaryColor: Colors.blue,
        elevatedButtonTheme: ElevatedButtonThemeData(
          style: ElevatedButton.styleFrom(
            minimumSize: Size(200, 50), // 버튼 크기 설정
          ),
        ),
      ),
      home: SplashScreen(), // 초기 화면을 스플래시 화면으로 설정
    );
  }
}

class MyHomePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Fast Deal'),
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            ElevatedButton(
              onPressed: () {
                Navigator.push(
                  context,
                  MaterialPageRoute(builder: (context) => ReportFeature()),
                );
              },
              child: Text('엑셀 기능'),
            ),
            SizedBox(height: 20), // 버튼 간 간격
            ElevatedButton(
              onPressed: () {
                Navigator.push(
                  context,
                  MaterialPageRoute(builder: (context) => ContractFeature()),
                );
              },
              child: Text('계약서 서명 기능'),
            ),
          ],
        ),
      ),
    );
  }
}
