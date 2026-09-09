// Basic tests for the PDIoT recording app.
//
// To perform an interaction with a widget in your test, use the WidgetTester
// utility in the flutter_test package.

import 'package:flutter_test/flutter_test.dart';

import 'package:pdiot/main.dart';
import 'package:pdiot/utils.dart';

void main() {
  testWidgets('Home page shows the recording controls',
      (WidgetTester tester) async {
    // Build our app and trigger a frame.
    await tester.pumpWidget(const MyApp());

    expect(find.text('PDIoT'), findsOneWidget);
    expect(find.text('Connect'), findsOneWidget);
    expect(find.text('Start recording'), findsOneWidget);
    expect(find.text('Stop recording'), findsOneWidget);
    expect(find.text('Settings'), findsOneWidget);
    expect(find.text('Not recording'), findsOneWidget);
  });

  group('sentenceToCamelCase', () {
    test('joins the words of an activity name', () {
      expect(sentenceToCamelCase('Lying down on left'), 'lyingDownOnLeft');
      expect(sentenceToCamelCase('Standing'), 'standing');
    });

    test('tolerates empty and oddly spaced input', () {
      expect(sentenceToCamelCase(''), '');
      expect(sentenceToCamelCase('   '), '');
      expect(sentenceToCamelCase('  Normal   walking '), 'normalWalking');
    });
  });

  group('combineAccelBytes', () {
    test('decodes positive values', () {
      expect(combineAccelBytes(0, 0), 0.0);
      // 0x4000 = 16384 = 1g
      expect(combineAccelBytes(0x40, 0x00), 1.0);
    });

    test('decodes negative values', () {
      // 0xC000 as a signed 16 bit value is -16384 = -1g. The upper byte is
      // read from the packet with getInt8, so it arrives here as -64.
      expect(combineAccelBytes(-64, 0x00), -1.0);
      // 0xFFFF = -1
      expect(combineAccelBytes(-1, 0xFF), -1 / 16384.0);
    });
  });
}
