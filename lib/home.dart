import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'settings.dart';
import 'utils.dart';
import 'ble.dart';
import 'globals.dart';
import 'dart:io';

// The home page includes buttons to connect to the respeck, enter capture
// metadata and start recording.

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  // This widget is the home page of your application. It is stateful, meaning
  // that it has a State object (defined below) that contains fields that affect
  // how it looks.

  // This class is the configuration for the state. It holds the values (in this
  // case the title) provided by the parent (in this case the App widget) and
  // used by the build method of the State. Fields in a Widget subclass are
  // always marked "final".

  final String title;

  @override
  State<MyHomePage> createState() => MyHomePageState();
}

class MyHomePageState extends State<MyHomePage> {
  //String _counter = "---";
  String accel = "";
  String batt_level = "";
  String recording_info = "Not recording";

  var activities = [
    'Standing',
    'Lying down on left',
    'Lying down right',
    'Lying down back',
    'Lying down on stomach',
    'Normal walking',
    'Ascending stairs',
    'Descending stairs',
    'Shuffle walking',
    'Running',
    'Miscellaneous movements'
  ];
  String selected_activity = "Standing";

  var social_signals = [
    'Normal',
    'Coughing',
    'Hyperventilating',
    'Talking',
    'Eating',
    'Singing',
    'Laughing'
  ];
  String selected_signal = "Normal";

  bool recording = false;
  bool received_packet = false;

  int recorded_samples = 0;
  DateTime? start_timestamp;

  String filename = "";

  // The UI for the main screen of the app is defined below. Unlike traditional
  // android code, there is no additional XML layout file.

  @override
  Widget build(BuildContext context) {
    // This method is rerun every time setState is called to update the UI.
    //
    // The Flutter framework has been optimized to make rerunning build methods
    // fast, so that you can just rebuild anything that needs updating rather
    // than having to individually change instances of widgets.
    return Scaffold(
      appBar: AppBar(
        // TRY THIS: Try changing the color here to a specific color (to
        // Colors.amber, perhaps?) and trigger a hot reload to see the AppBar
        // change color while the other colors stay the same.
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        // Here we take the value from the MyHomePage object that was created by
        // the App.build method, and use it to set our appbar title.
        title: Text(widget.title),
      ),
      body: Center(
        // Center is a layout widget. It takes a single child and positions it
        // in the middle of the parent.
        child: Column(
          // Column is also a layout widget. It takes a list of children and
          // arranges them vertically. By default, it sizes itself to fit its
          // children horizontally, and tries to be as tall as its parent.
          //
          // Column has various properties to control how it sizes itself and
          // how it positions its children. Here we use mainAxisAlignment to
          // center the children vertically; the main axis here is the vertical
          // axis because Columns are vertical (the cross axis would be
          // horizontal).
          //
          // TRY THIS: Invoke "debug painting" (choose the "Toggle Debug Paint"
          // action in the IDE, or press "p" in the console), to see the
          // wireframe for each widget.
          mainAxisAlignment: MainAxisAlignment.start,
          children: <Widget>[
            const SizedBox(height: 10),
            ElevatedButton(
                onPressed: () async {
                  if (respeckUUID == null || respeckUUID == "") {
                    showToast("Please pair with a Respeck first");
                    return;
                  }
                  await scanForRespeck(this);
                  await connectToRespeck();
                  scanForRespeck(this);
                },
                style:
                    ElevatedButton.styleFrom(backgroundColor: Colors.lightBlue),
                child: const Text('Connect')),
            const SizedBox(height: 10),
            const Text(
              'Acceleration (g)',
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            Text(
              accel,
            ),
            Text(
              batt_level,
            ),
            const SizedBox(height: 10),
            const Text(
              'Activity',
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            DropdownButton(
              // Initial Value
              value: selected_activity,
              items: activities.map((String items) {
                return DropdownMenuItem(value: items, child: Text(items));
              }).toList(),
              // After selecting the desired option,it will
              // change button value to selected value
              onChanged: (String? newValue) {
                setState(() {
                  selected_activity = newValue!;
                });
              },
            ),
            const SizedBox(height: 10),
            const Text(
              'Social signal',
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            DropdownButton(
              // Initial Value
              value: selected_signal,
              items: social_signals.map((String items) {
                return DropdownMenuItem(value: items, child: Text(items));
              }).toList(),
              // After selecting the desired option,it will
              // change button value to selected value
              onChanged: (String? newValue) {
                setState(() {
                  selected_signal = newValue!;
                });
              },
            ),
            const SizedBox(height: 10),
            ElevatedButton(
                onPressed: record,
                style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.lightGreen),
                child: const Text('Start recording')),
            const SizedBox(height: 10),
            const Text(
              "Recording status",
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            Text(
              recording_info,
            ),
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 10.0),
              child: Text(
                filename,
              ),
            ),
            const SizedBox(height: 10),
            ElevatedButton(
                onPressed: () {
                  if (!recording) {
                    return;
                  }
                  recording = false;
                  showToast("Recording stopped");
                },
                style:
                    ElevatedButton.styleFrom(backgroundColor: Colors.red[200]!),
                child: const Text('Stop recording')),
            const SizedBox(height: 10),
            ElevatedButton(
                onPressed: () {
                  Navigator.push(
                    context,
                    MaterialPageRoute(
                        builder: (context) => const SettingsPage()),
                  );
                },
                style: ElevatedButton.styleFrom(backgroundColor: Colors.grey),
                child: const Text('Settings')),
          ],
        ),
      ),
    );
  }

  // Start recording respeck data to CSV
  void record() async {
    if (!received_packet) {
      showToast("Please connect to a Respeck first");
      return;
    }
    if (recording) {
      showToast("Already recording");
      return;
    }

    recorded_samples = 0;
    DateTime now =
        DateTime.now().toUtc(); //use current UTC timestamp for filename
    start_timestamp = now;
    String formattedDate =
        '${DateFormat('yyyy-MM-dd').format(now)}T${DateFormat('kkmmss').format(now)}Z';
    filename =
        'PDIoT_${subjectID}_${sentenceToCamelCase(selected_activity)}_${sentenceToCamelCase(selected_signal)}_${formattedDate}_${respeckUUID?.replaceAll(":", "")}.csv';
    print(filename);

    // create file and write CSV header row
    csvFile = File('${storageFolder?.path}/$filename');
    await csvFile?.writeAsString(
        "receivedPhoneTimestamp,respeckTimestamp,packetSeqNum,sampleSeqNum,accelX,accelY,accelY\n",
        flush: true);
    showToast("Recording started..");
    recording = true;
  }
}
