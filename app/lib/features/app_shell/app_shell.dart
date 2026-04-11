import 'package:flutter/material.dart';

import '../../core/localization/app_strings.dart';
import '../capture/capture_hub_screen.dart';
import '../home/home_screen.dart';
import '../library/library_screen.dart';
import '../report_packs/report_packs_screen.dart';
import '../settings/settings_screen.dart';

class AppShell extends StatefulWidget {
  final int initialIndex;

  const AppShell({super.key, this.initialIndex = 0});

  @override
  State<AppShell> createState() => _AppShellState();
}

class _AppShellState extends State<AppShell> {
  late int _index = widget.initialIndex;

  late final List<Widget> _screens = [
    const HomeScreen(),
    const LibraryScreen(),
    const CaptureHubScreen(),
    const ReportPacksScreen(),
    const SettingsScreen(),
  ];

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    return Scaffold(
      body: IndexedStack(index: _index, children: _screens),
      bottomNavigationBar: NavigationBar(
        selectedIndex: _index,
        onDestinationSelected: (value) => setState(() => _index = value),
        destinations: [
          NavigationDestination(icon: const Icon(Icons.home_outlined), selectedIcon: const Icon(Icons.home), label: s.home),
          NavigationDestination(icon: const Icon(Icons.folder_open_outlined), selectedIcon: const Icon(Icons.folder_open), label: s.library),
          NavigationDestination(icon: const Icon(Icons.add_box_outlined), selectedIcon: const Icon(Icons.add_box), label: s.capture),
          NavigationDestination(icon: const Icon(Icons.picture_as_pdf_outlined), selectedIcon: const Icon(Icons.picture_as_pdf), label: s.reports),
          NavigationDestination(icon: const Icon(Icons.settings_outlined), selectedIcon: const Icon(Icons.settings), label: s.settings),
        ],
      ),
    );
  }
}
