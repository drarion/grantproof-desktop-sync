import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../core/utils/date_utils.dart';
import '../../shared/widgets/empty_state.dart';
import '../../shared/widgets/frosted_card.dart';
import 'workspace_connection_screen.dart';

class SyncCenterScreen extends StatelessWidget {
  const SyncCenterScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final pendingEvidence = state.evidences.where((item) => !item.isSynced).toList();
    final pendingStories = state.stories.where((item) => !item.isSynced).toList();

    String providerSummary() {
      if (!state.workspaceConnected) {
        return s.connectWorkspaceFirst;
      }
      if (state.workspaceProvider == 'desktop') {
        if (state.desktopPairingReady) {
          return s.isFr
              ? 'Desktop Sync actif • ${state.desktopComputerLabel} • ${state.desktopServerHost}:${state.desktopServerPort}'
              : 'Desktop Sync active • ${state.desktopComputerLabel} • ${state.desktopServerHost}:${state.desktopServerPort}';
        }
        return s.isFr
            ? 'Desktop Sync configuré pour ${state.workspaceName}. Il faut encore jumeler ce téléphone avec le QR du compagnon PC.'
            : 'Desktop Sync is configured for ${state.workspaceName}. You still need to pair this phone with the desktop companion QR.';
      }
      if (state.workspaceNeedsAccountConnection) {
        return s.isFr
            ? '${state.workspaceProviderLabel} importé : connecte maintenant le compte utilisé sur cet appareil.'
            : '${state.workspaceProviderLabel} imported: now connect the account used on this device.';
      }
      return '${state.workspaceProviderLabel} • ${state.workspaceName}';
    }

    return Scaffold(
      appBar: AppBar(title: Text(s.syncCenter)),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 40),
        children: [
          Text(s.syncSubtitle, style: Theme.of(context).textTheme.bodyLarge),
          const SizedBox(height: 20),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Wrap(
                  spacing: 10,
                  runSpacing: 10,
                  children: [
                    Chip(label: Text('${s.evidencePending}: ${pendingEvidence.length}')),
                    Chip(label: Text('${s.storiesPending}: ${pendingStories.length}')),
                    Chip(label: Text('${s.lastSync}: ${state.lastSyncedAt == null ? '—' : DateUtilsX.formatShort(state.lastSyncedAt!)}')),
                    Chip(
                      label: Text(
                        !state.workspaceConnected
                            ? s.workspaceNotConnected
                            : state.workspaceNeedsAccountConnection
                                ? (s.isFr ? 'Compte requis' : 'Account required')
                                : (state.workspaceProvider == 'desktop'
                                ? (state.desktopPairingReady ? 'Desktop Sync' : (s.isFr ? 'Jumelage requis' : 'Pairing required'))
                                : s.workspaceConnected),
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 14),
                Text(providerSummary()),
                const SizedBox(height: 16),
                Wrap(
                  spacing: 12,
                  runSpacing: 12,
                  children: [
                    FilledButton.icon(
                      onPressed: state.workspaceCanSyncNow || state.isSyncing
                          ? () async {
                              final messenger = ScaffoldMessenger.of(context);
                              try {
                                final count = await context.read<AppState>().syncNow();
                                if (context.mounted) {
                                  final message = state.workspaceProvider == 'desktop'
                                      ? (s.isFr
                                          ? 'Synchronisation locale vers ${state.desktopComputerLabel} terminée : $count élément(s).'
                                          : 'Local sync to ${state.desktopComputerLabel} complete: $count item(s).')
                                      : (s.isFr
                                          ? 'Synchronisation Google Drive terminée : $count élément(s).'
                                          : 'Google Drive sync complete: $count item(s).');
                                  messenger.showSnackBar(SnackBar(content: Text(message)));
                                }
                              } catch (error) {
                                if (context.mounted) {
                                  messenger.showSnackBar(
                                    SnackBar(
                                      content: Text(
                                        s.isFr ? 'Erreur de synchronisation : $error' : 'Sync failed: $error',
                                      ),
                                    ),
                                  );
                                }
                              }
                            }
                          : null,
                      icon: Icon(state.isSyncing ? Icons.sync : Icons.cloud_upload_outlined),
                      label: Text(state.isSyncing ? s.syncInProgress : s.syncNow),
                    ),
                    OutlinedButton.icon(
                      onPressed: () => Navigator.of(context).push(
                        MaterialPageRoute<void>(builder: (_) => const WorkspaceConnectionScreen()),
                      ),
                      icon: const Icon(Icons.tune_rounded),
                      label: Text(s.openWorkspace),
                    ),
                    if (state.workspaceProvider == 'desktop' && state.organizationProvisioningCode.isNotEmpty)
                      OutlinedButton.icon(
                        onPressed: () async {
                          await Clipboard.setData(ClipboardData(text: state.organizationProvisioningCode));
                          if (context.mounted) {
                            ScaffoldMessenger.of(context).showSnackBar(
                              SnackBar(
                                content: Text(
                                  s.isFr ? 'Code d’organisation copié.' : 'Organization code copied.',
                                ),
                              ),
                            );
                          }
                        },
                        icon: const Icon(Icons.copy_all_outlined),
                        label: Text(s.isFr ? 'Copier le code' : 'Copy code'),
                      ),
                  ],
                ),
                if (state.workspaceProvider == 'desktop') ...[
                  const SizedBox(height: 12),
                  Text(
                    state.desktopPairingReady
                        ? (s.isFr
                            ? 'Les preuves seront déposées dans le dossier ${state.desktopFolderLabel} sur ${state.desktopComputerLabel.isEmpty ? 'votre ordinateur' : state.desktopComputerLabel}. Les autres fonctions de collecte restent inchangées.'
                            : 'Evidence will be written into the ${state.desktopFolderLabel} folder on ${state.desktopComputerLabel.isEmpty ? 'your computer' : state.desktopComputerLabel}. All capture functions remain unchanged.')
                        : (s.isFr
                            ? 'Desktop Sync est configuré mais ce téléphone doit encore être jumelé avant le premier envoi local.'
                            : 'Desktop Sync is configured but this phone still needs pairing before the first local transfer.'),
                    style: Theme.of(context).textTheme.bodySmall,
                  ),
                ],
              ],
            ),
          ),
          const SizedBox(height: 20),
          if (pendingEvidence.isEmpty && pendingStories.isEmpty)
            EmptyState(
              icon: Icons.cloud_done_outlined,
              iconColor: AppTheme.success,
              title: s.noPendingSync,
              subtitle: s.noPendingSyncSubtitle,
            )
          else ...[
            ...pendingEvidence.map(
              (item) => Padding(
                padding: const EdgeInsets.only(bottom: 12),
                child: _SyncCard(
                  icon: Icons.verified_outlined,
                  title: item.title,
                  subtitle: [item.activity, item.locationLabel].where((e) => e.trim().isNotEmpty).join(' • '),
                ),
              ),
            ),
            ...pendingStories.map(
              (item) => Padding(
                padding: const EdgeInsets.only(bottom: 12),
                child: _SyncCard(
                  icon: Icons.menu_book_outlined,
                  title: item.title,
                  subtitle: item.beneficiaryAlias,
                ),
              ),
            ),
          ],
        ],
      ),
    );
  }
}

class _SyncCard extends StatelessWidget {
  final IconData icon;
  final String title;
  final String subtitle;

  const _SyncCard({required this.icon, required this.title, required this.subtitle});

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    return FrostedCard(
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(
            width: 44,
            height: 44,
            decoration: BoxDecoration(
              color: AppTheme.primary.withValues(alpha: 0.08),
              borderRadius: BorderRadius.circular(16),
            ),
            child: Icon(icon, color: AppTheme.primary),
          ),
          const SizedBox(width: 14),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  title,
                  style: Theme.of(context).textTheme.titleMedium,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                ),
                if (subtitle.trim().isNotEmpty) ...[
                  const SizedBox(height: 6),
                  Text(
                    subtitle,
                    style: Theme.of(context).textTheme.bodyMedium,
                    maxLines: 2,
                    overflow: TextOverflow.ellipsis,
                  ),
                ],
                const SizedBox(height: 14),
                Align(
                  alignment: Alignment.centerLeft,
                  child: Chip(label: Text(s.pendingSync)),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
