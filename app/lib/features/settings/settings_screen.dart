import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../shared/widgets/frosted_card.dart';
import '../workspace/sync_center_screen.dart';
import '../workspace/workspace_connection_screen.dart';

class SettingsScreen extends StatelessWidget {
  const SettingsScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    return SafeArea(
      child: ListView(
        padding: const EdgeInsets.fromLTRB(20, 16, 20, 120),
        children: [
          Text(s.settings, style: Theme.of(context).textTheme.headlineLarge),
          const SizedBox(height: 8),
          Text(s.settingsSubtitle, style: Theme.of(context).textTheme.bodyLarge),
          const SizedBox(height: 24),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.workspace, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 10),
                Text(state.workspaceConnected ? state.workspaceName : s.workspaceNotConnected),
                const SizedBox(height: 6),
                Text(!state.workspaceConnected ? s.disconnectedNotice : state.workspaceNeedsAccountConnection ? (s.isFr ? '${state.workspaceProviderLabel} configuré • compte local à connecter' : '${state.workspaceProviderLabel} configured • local account still needs connection') : '${state.workspaceProviderLabel} • ${state.workspaceLibrary}'),
                const SizedBox(height: 16),
                Wrap(
                  spacing: 12,
                  runSpacing: 12,
                  children: [
                    FilledButton.tonalIcon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => const WorkspaceConnectionScreen()),
                      ),
                      icon: const Icon(Icons.link_outlined),
                      label: Text(s.openWorkspace),
                    ),
                    OutlinedButton.icon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => const SyncCenterScreen()),
                      ),
                      icon: const Icon(Icons.sync_outlined),
                      label: Text(s.openSyncCenter),
                    ),
                  ],
                ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.language, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 12),
                SegmentedButton<String>(
                  segments: const [
                    ButtonSegment(value: 'fr', label: Text('FR')),
                    ButtonSegment(value: 'en', label: Text('ENG')),
                  ],
                  selected: <String>{state.languageCode},
                  onSelectionChanged: (values) => context.read<AppState>().setLanguage(values.first),
                ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.demoWorkspace, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 10),
                Text(s.isFr ? 'Mode principal recommandé : GrantProof Desktop Sync' : 'Recommended primary mode: GrantProof Desktop Sync'),
                const SizedBox(height: 6),
                Text(s.isFr ? 'Alternative avancée : Google Workspace ou Microsoft 365 selon les contraintes IT du client.' : 'Advanced alternative: Google Workspace or Microsoft 365 depending on the client IT constraints.'),
                const SizedBox(height: 6),
                Text(s.isFr ? 'Aucun stockage GrantProof requis : les preuves restent chez le client.' : 'No GrantProof storage required: evidence stays with the client.'),
              ],
            ),
          ),
          const SizedBox(height: 16),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(s.utilities, style: Theme.of(context).textTheme.titleMedium),
                const SizedBox(height: 12),
                OutlinedButton.icon(
                  onPressed: () async {
                    await context.read<AppState>().resetDemo();
                    if (context.mounted) {
                      ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(s.demoRestored)));
                    }
                  },
                  icon: const Icon(Icons.refresh_outlined),
                  label: Text(s.resetDemoData),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
