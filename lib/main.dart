import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'dart:io';
import 'globals.dart';
import 'package:package_info_plus/package_info_plus.dart';
import 'home.dart';

// This file starts the app and displays the home screen.

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  // Read app version information. This must be awaited, otherwise the settings
  // page can be built before the real version has arrived and will show the
  // "0.0.0" default instead.
  try {
    final PackageInfo packageInfo = await PackageInfo.fromPlatform();
    appVersionName = packageInfo.version; // App version
    appVersionCode = int.tryParse(packageInfo.buildNumber) ?? 0; // Build number
  } catch (e) {
    // A missing version is not worth blocking startup for
    print("Could not read package info: $e");
  }

  // Read pairing information form shared preferences
  respeckUUID = await asyncPrefs.getString('rid');
  print("respeckUUID:${respeckUUID}");

  subjectID = await asyncPrefs.getString('sid');
  print("subjectID:${subjectID}");

  // get storage folder - must be accessible to the user for PDIoT
  try {
    if (Platform.isAndroid) {
      storageFolder = await getDownloadsDirectory();
    } else {
      storageFolder = await getApplicationDocumentsDirectory();
    }
  } catch (e) {
    print("Could not get storage folder: $e");
  }

  // Fall back to the app documents directory if the downloads folder was not
  // available, so that recording still works (the files are just harder to
  // reach over USB)
  if (storageFolder == null) {
    try {
      storageFolder = await getApplicationDocumentsDirectory();
    } catch (e) {
      print("Could not get fallback storage folder: $e");
    }
  }
  print("storageFolder:${storageFolder?.path}");

  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'pdiot',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.lightBlue),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: "PDIoT"),
    );
  }
}
