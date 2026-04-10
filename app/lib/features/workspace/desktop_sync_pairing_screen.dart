import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/services/desktop_sync_service.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../shared/widgets/frosted_card.dart';
import '../app_shell/app_shell.dart';
import 'provisioning_scanner_screen.dart';

class DesktopSyncPairingScreen extends StatefulWidget {
  const DesktopSyncPairingScreen({super.key});

  @override
  State<DesktopSyncPairingScreen> createState() => _DesktopSyncPairingScreenState();
}

class _DesktopSyncPairingScreenState extends State<DesktopSyncPairingScreen> {
  Future<void> _goHome() async {
    if (!mounted) return;
    Navigator.of(context).pushAndRemoveUntil(
      MaterialPageRoute<void>(builder: (_) => const AppShell(initialIndex: 0)),
      (route) => false,
    );
  }

  Future<void> _pairFromPayload(String rawPayload) async {
    final s = AppStrings.of(context);
    final payload = rawPayload.trim();
    if (payload.isEmpty) return;

    try {
      if (!DesktopSyncService.instance.looksLikePairingPayload(payload)) {
        throw FormatException(
          s.isFr
              ? 'Ce code ne correspond pas au QR de jumelage du compagnon PC.'
              : 'This code does not match the pairing QR from the desktop companion.',
        );
      }
      await context.read<AppState>().pairDesktopCompanion(payload);
      if (!mounted) return;
      final state = context.read<AppState>();
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            s.isFr
                ? 'Ordinateur jumelé : ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}'
                : 'Computer paired: ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}',
          ),
        ),
      );
      setState(() {});
      await showDialog<void>(
        context: context,
        builder: (context) => AlertDialog(
          title: Text(s.isFr ? 'Connexion établie' : 'Connection established'),
          content: Text(
            s.isFr
                ? 'Le téléphone est maintenant jumelé au compagnon PC. Vous pouvez revenir directement à l’accueil.'
                : 'The phone is now paired with the desktop companion. You can go straight back to the home screen.',
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: Text(s.isFr ? 'Rester ici' : 'Stay here'),
            ),
            FilledButton(
              onPressed: () {
                Navigator.of(context).pop();
                _goHome();
              },
              child: Text(s.returnHome),
            ),
          ],
        ),
      );
    } catch (error) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            s.isFr ? 'Jumelage impossible : $error' : 'Pairing failed: $error',
          ),
        ),
      );
    }
  }

  Future<void> _scanPcQr() async {
    final raw = await Navigator.of(context).push<String>(
      MaterialPageRoute<String>(
        builder: (_) => ProvisioningScannerScreen(
          title: AppStrings.of(context).isFr ? 'Scanner le QR du PC' : 'Scan the PC QR',
        ),
      ),
    );
    if (!mounted || raw == null) return;
    await _pairFromPayload(raw);
  }

  Future<void> _pastePairingCode() async {
    final code = await showModalBottomSheet<String>(
      context: context,
      isScrollControlled: true,
      showDragHandle: true,
      builder: (context) => const _PairingCodeSheet(),
    );
    if (!mounted || code == null) return;
    await _pairFromPayload(code);
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final state = context.watch<AppState>();

    return Scaffold(
      appBar: AppBar(
        title: Text(s.isFr ? 'Jumeler le téléphone' : 'Pair the phone'),
        actions: [
          if (state.desktopPairingReady)
            IconButton(
              onPressed: _goHome,
              tooltip: s.returnHome,
              icon: const Icon(Icons.home_outlined),
            ),
        ],
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 32),
        children: [
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Container(
                      width: 48,
                      height: 48,
                      decoration: BoxDecoration(
                        color: state.desktopPairingReady
                            ? AppTheme.success.withValues(alpha: 0.10)
                            : AppTheme.primary.withValues(alpha: 0.08),
                        borderRadius: BorderRadius.circular(16),
                      ),
                      child: Icon(
                        Icons.qr_code_rounded,
                        color: state.desktopPairingReady ? AppTheme.success : AppTheme.primary,
                      ),
                    ),
                    const SizedBox(width: 12),
                    Expanded(
                      child: Text(
                        s.isFr ? 'Étape 3 · Jumeler avec le QR du PC' : 'Step 3 · Pair with the PC QR',
                        style: Theme.of(context).textTheme.titleMedium,
                      ),
                    ),
                    Chip(
                      label: Text(
                        state.desktopPairingReady
                            ? (s.isFr ? 'Jumelé' : 'Paired')
                            : (s.isFr ? 'À faire' : 'To do'),
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 12),
                Text(
                  state.desktopPairingReady
                      ? (s.isFr
                          ? 'Ce téléphone est déjà relié à ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}. Vous pouvez synchroniser vers le dossier ${state.desktopFolderLabel}.'
                          : 'This phone is already linked to ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}. You can sync into the ${state.desktopFolderLabel} folder.')
                      : (s.isFr
                          ? 'Ouvrez GrantProof Desktop Sync sur le PC cible, affichez le QR de jumelage puis scannez-le ici.'
                          : 'Open GrantProof Desktop Sync on the target PC, display the pairing QR, then scan it here.'),
                ),
              ],
            ),
          ),
          const SizedBox(height: 18),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  s.isFr ? 'Configuration en cours' : 'Current setup',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 10),
                Text(
                  s.isFr ? 'Organisation : ${state.workspaceName}' : 'Organization: ${state.workspaceName}',
                ),
                const SizedBox(height: 4),
                Text(
                  s.isFr
                      ? 'Ordinateur cible : ${state.desktopComputerLabel.isEmpty ? 'À renseigner' : state.desktopComputerLabel}'
                      : 'Target computer: ${state.desktopComputerLabel.isEmpty ? 'To be set' : state.desktopComputerLabel}',
                ),
                const SizedBox(height: 4),
                Text(
                  s.isFr ? 'Dossier local : ${state.desktopFolderLabel}' : 'Local folder: ${state.desktopFolderLabel}',
                ),
                if (state.desktopPairingReady) ...[
                  const SizedBox(height: 4),
                  Text(
                    s.isFr
                        ? 'Connexion locale : ${state.desktopServerHost}:${state.desktopServerPort}'
                        : 'Local connection: ${state.desktopServerHost}:${state.desktopServerPort}',
                  ),
                ],
              ],
            ),
          ),
          const SizedBox(height: 18),
          Wrap(
            spacing: 10,
            runSpacing: 10,
            children: [
              FilledButton.icon(
                onPressed: _scanPcQr,
                icon: const Icon(Icons.qr_code_scanner_rounded),
                label: Text(s.isFr ? 'Scanner le QR du PC' : 'Scan the PC QR'),
              ),
              OutlinedButton.icon(
                onPressed: _pastePairingCode,
                icon: const Icon(Icons.content_paste_go_rounded),
                label: Text(s.isFr ? 'Coller un code de jumelage' : 'Paste a pairing code'),
              ),
              if (state.desktopPairingReady)
                OutlinedButton.icon(
                  onPressed: () => context.read<AppState>().clearDesktopPairing(),
                  icon: const Icon(Icons.link_off_rounded),
                  label: Text(s.isFr ? 'Oublier ce jumelage' : 'Forget this pairing'),
                ),
            ],
          ),
          if (state.desktopPairingReady) ...[
            const SizedBox(height: 18),
            FilledButton.icon(
              onPressed: _goHome,
              icon: const Icon(Icons.home_outlined),
              label: Text(s.returnHome),
            ),
          ],
        ],
      ),
    );
  }
}

class _PairingCodeSheet extends StatefulWidget {
  const _PairingCodeSheet();

  @override
  State<_PairingCodeSheet> createState() => _PairingCodeSheetState();
}

class _PairingCodeSheetState extends State<_PairingCodeSheet> {
  final TextEditingController _codeController = TextEditingController();

  @override
  void dispose() {
    _codeController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final bottom = MediaQuery.of(context).viewInsets.bottom;

    return Padding(
      padding: EdgeInsets.fromLTRB(20, 8, 20, bottom + 24),
      child: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            s.isFr ? 'Coller un code de jumelage' : 'Paste a pairing code',
            style: Theme.of(context).textTheme.titleLarge,
          ),
          const SizedBox(height: 8),
          Text(
            s.isFr
                ? 'Collez ici le code brut affiché par GrantProof Desktop Sync si vous ne passez pas par le scanner QR.'
                : 'Paste here the raw code displayed by GrantProof Desktop Sync if you are not using the QR scanner.',
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _codeController,
            minLines: 4,
            maxLines: 8,
            decoration: InputDecoration(
              hintText: s.isFr ? 'grantproof://pair?...' : 'grantproof://pair?...',
            ),
          ),
          const SizedBox(height: 16),
          FilledButton(
            onPressed: () => Navigator.of(context).pop(_codeController.text.trim()),
            child: Text(s.isFr ? 'Utiliser ce code' : 'Use this code'),
          ),
        ],
      ),
    );
  }
}
