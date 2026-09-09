import 'package:flutter_blue_plus/flutter_blue_plus.dart';
import 'dart:io';
import 'dart:typed_data';
import 'utils.dart';
import 'home.dart';
import 'globals.dart';

// This file contains the BLE code to scan for and connect to a respeck sensor.
// Once connected, we notify the acceleration characteristic and decode the
// stream of packets to obtain x,y,z acceleration in g and battery level.
// Each sample is written to a CSV file and the UI is updated with the latest
// values.

BluetoothDevice? respeck;
late MyHomePageState ui;
var subscription2;
late int respeckVersion;

Future<void> scanForRespeck(MyHomePageState _ui) async {
  ui = _ui;
  // listen to scan results
  // Note: `onScanResults` clears the results between scans. You should use
  //  `scanResults` if you want the current scan results *or* the results from the previous scan.

  showToast("Searching for Respeck $respeckUUID...");

  var subscription = FlutterBluePlus.scanResults.listen(
    (results) {
      if (results.isNotEmpty) {
        for (ScanResult r in results) {
          if (Platform.isAndroid) {
            //ScanResult r = results.last; // the most recently found device
            //print('${r.device.remoteId}: "${r.advertisementData.advName}" found!');
            if (r.device.remoteId.str == respeckUUID) {
              if (r.advertisementData.advName == "Res6AL") {
                respeck = r.device;
                respeckVersion = 6;
                fwString = "6AL";
              } else if (r.advertisementData.advName == "Res6AM") {
                respeck = r.device;
                respeckVersion = 6;
                fwString = "6AM";
              } else if (r.advertisementData.advName == "ResV5i") {
                respeck = r.device;
                respeckVersion = 5;
                fwString = "5i";
              } else {
                showToast("Respeck firmware is too old");
              }
            }
          } else if (Platform.isIOS) {
            //ScanResult r = results.last; // the most recently found device
            //print('${r.device.remoteId}: "${r.advertisementData.advName}" found!');
            final serviceData =
                r.advertisementData.serviceData; // Map<Guid, List<int>>
            final feedUuid = Guid('0000feed-0000-1000-8000-00805f9b34fb');
            final bytes = serviceData[feedUuid];
            if (bytes != null && bytes.isNotEmpty) {
              final hex = bytes
                  .map((b) => b.toRadixString(16).padLeft(2, '0'))
                  .toList();
              final hex_str = hex.join(':').toUpperCase();
              //print('Found service data: GAP UUID: $hex_str');
              if (r.advertisementData.advName == "Res6AM") {
                if (hex_str == respeckUUID) {
                  respeck = r.device;
                  respeckVersion = 6;
                  fwString = "6AM";
                }
              } else {
                showToast("Respeck firmware is too old");
              }
            }
          }
          FlutterBluePlus.stopScan();
          print("STOP SCANNING");
          break;
        }
      }
    },
    onError: (e) => print(e),
  );

  // cleanup: cancel subscription when scanning stops
  FlutterBluePlus.cancelWhenScanComplete(subscription);

  // Wait for Bluetooth enabled & permission granted
  // In your real app you should use `FlutterBluePlus.adapterState.listen` to handle all states
  await FlutterBluePlus.adapterState
      .where((val) => val == BluetoothAdapterState.on)
      .first;

  // Start scanning w/ timeout
  // Optional: use `stopScan()` as an alternative to timeout
  await FlutterBluePlus.startScan(
      //withServices: [Guid("180D")], // match any of the specified services
      //withNames: ["Bluno"], // *or* any of the specified names
      //timeout: const Duration(seconds: 5)
      );

  // wait for scanning to stop
  await FlutterBluePlus.isScanning.where((val) => val == false).first;
  print("end of scanning");
}

Future<void> notify() async {
  BluetoothService? ser;
  BluetoothCharacteristic? cha;

  // Service discovery
  // Note: You must call discoverServices after every re-connection!
  List<BluetoothService> services = await respeck!.discoverServices();
  //services.forEach((service) {
  for (BluetoothService s in services) {
    if (s.serviceUuid.str == "00001523-1212-efde-1523-785feabcd125") {
      ser = s;
      break;
    }
    // Reads all characteristics
  }

  // Read all characteristics provided by this service and choose the
  // acceleration characteristic

  if (ser != null) {
    var characteristics = ser.characteristics;
    for (BluetoothCharacteristic c in characteristics) {
      if (c.characteristicUuid.str == "00001524-1212-efde-1523-785feabcd125") {
        print("found accel characteristic");
        cha = c;
        break;
      }
    }
  }

  // Subscribe to acceleration values
  if (cha != null) {
    final subscription3 = cha.onValueReceived.listen((value) async {
      ui.received_packet = true;
      // onValueReceived is updated:
      //   - anytime read() is called
      //   - anytime a notification arrives (if subscribed)

      //print("RECEIVED packet:");
      //print(value);

      // Decode packet:
      //https://stackoverflow.com/questions/59627795/parse-int-and-float-values-from-uint8list-dart
      //https://stackoverflow.com/questions/72418830/how-to-target-specific-int-bits-in-an-int-array-into-uint16-or-32

      // Timestamp when this packet was received, according to the phone clock
      DateTime packet_received_ts = DateTime.now();

      Uint8List ul = Uint8List.fromList(value);
      ByteData bd = ul.buffer.asByteData();

      int ts = bd.getUint32(0);
      ts = (ts * 197 * 1000 / 32768)
          .toInt(); // timestamp from the respeck clock, in ms
      //print(ts);

      // Sequence number can be used to detect dropped packets
      int packetSeqNumber = bd.getUint16(4);
      //print(packetSeqNumber);

      // Battery status (respeck version 6 only)
      int? battLevel;
      bool charging = false;

      if (respeckVersion == 6) {
        battLevel = bd.getUint8(6);

        if (bd.getUint8(7) == 1) {
          charging = true;
        }
      }

      // NOTE: The respeck sends multiple samples per packet, for efficiency
      // For each sample in this packet, append a line to the CSV file,

      String csv_str = "";
      int seqNumInPacket = 0;
      double x = 0, y = 0, z = 0;

      for (int i = 8; i < bd.lengthInBytes; i += 6) {
        int b1 = bd.getInt8(i);
        int b2 = bd.getInt8(i + 1);
        x = combineAccelBytes(b1, b2);
        int b3 = bd.getInt8(i + 2);
        int b4 = bd.getInt8(i + 3);
        y = combineAccelBytes(b3, b4);
        int b5 = bd.getInt8(i + 4);
        int b6 = bd.getInt8(i + 5);
        z = combineAccelBytes(b5, b6);

        csv_str +=
            "${packet_received_ts.millisecondsSinceEpoch},$ts,$packetSeqNumber,$seqNumInPacket,$x,$y,$z\n";

        seqNumInPacket++;
        ui.recorded_samples++;

        // Update the UI to show the latest data (called once per packet)
        ui.setState(() {
          // This call to setState tells the Flutter framework that something has
          // changed in this State, which causes it to rerun the build method below
          // so that the display can reflect the updated values.
          ui.accel =
              "x=${x.toStringAsFixed(3)}, y=${y.toStringAsFixed(3)}, z=${z.toStringAsFixed(3)}";
          if (respeckVersion == 6) {
            if (charging) {
              ui.batt_level = "Battery: $battLevel% (charging)";
            } else {
              ui.batt_level = "Battery: $battLevel%";
            }
          } else {
            ui.batt_level = "";
          }
        });
      }
      // update elapsed time counter if recording
      if (ui.recording) {
        int elapsed_secs =
            packet_received_ts.difference(ui.start_timestamp!).inSeconds;

        ui.recording_info =
            "Written ${ui.recorded_samples} samples (${elapsed_secs} seconds)";

        // write to the CSV file
        await csvFile?.writeAsString(csv_str,
            mode: FileMode.writeOnlyAppend, flush: false);
      }
    });

    // cleanup: cancel subscription when disconnected
    respeck!.cancelWhenDisconnected(subscription3);

    await Future.delayed(const Duration(seconds: 3));

    // subscribe
    // Note: If a characteristic supports both **notifications** and **indications**,
    // it will default to **notifications**. This matches how CoreBluetooth works on iOS.
    await cha.setNotifyValue(true);
  }
}

Future<void> connectToRespeck() async {
  // listen for disconnection
  subscription2 =
      respeck!.connectionState.listen((BluetoothConnectionState state) async {
    if (state == BluetoothConnectionState.connected) {
      print("CONNECTED to Respeck $respeckUUID");
      respeckConnected = true;
      showToast("Connected to Respeck $respeckUUID");
      await notify();
    }

    if (state == BluetoothConnectionState.disconnected) {
      // 1. typically, start a periodic timer that tries to
      //    reconnect, or just call connect() again right now
      // 2. you must always re-discover services after disconnection!
      print(
          "DISCONNECT: ${respeck?.disconnectReason?.code} ${respeck?.disconnectReason?.description}");
      respeckConnected = false;

      if (respeck?.disconnectReason?.code != null &&
          respeck?.disconnectReason?.code != 0) {
        showLongToast("Waiting for reconnect...");
      }
    }
  });
  // cleanup: cancel subscription when disconnected
  //   - [delayed] This option is only meant for `connectionState` subscriptions.
  //     When `true`, we cancel after a small delay. This ensures the `connectionState`
  //     listener receives the `disconnected` event.
  //   - [next] if true, the the stream will be canceled only on the *next* disconnection,
  //     not the current disconnection. This is useful if you setup your subscriptions
  //     before you connect.
  //respeck!.cancelWhenDisconnected(subscription2, delayed: true, next: true);

  // Connect to the device
  if (respeck == null) {
    return;
  }

  // Now connect to the respeck
  // enable auto connect
  //  - note: autoConnect is incompatible with mtu argument, so you must call requestMtu yourself
  await respeck!
      .connect(license: License.nonprofit, autoConnect: true, mtu: null);

// wait until connection
//  - when using autoConnect, connect() always returns immediately, so we must
//    explicity listen to `device.connectionState` to know when connection occurs
//  await respeck!.connectionState
//      .where((val) => val == BluetoothConnectionState.connected)
//      .first;

//  print("CONNECTED to Respeck ${respeck!.remoteId}");
//  respeckConnected = true;
//  showToast("Connected to Respeck ${respeck!.remoteId}");

  //await notify();
}

Future<void> disconnect() async {
  await respeck!.disconnect();
  // cancel to prevent duplicate listeners
  await subscription2.cancel();
  csvFile = null;
}
