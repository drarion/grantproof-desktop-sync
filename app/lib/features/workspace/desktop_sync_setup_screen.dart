import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../shared/widgets/frosted_card.dart';
import 'desktop_sync_pairing_screen.dart';

class DesktopSyncSetupScreen extends StatefulWidget {
  const DesktopSyncSetupScreen({super.key});

  @override
  State<DesktopSyncSetupScreen> createState() => _DesktopSyncSetupScreenState();
}

class _DesktopSyncSetupScreenState extends State<DesktopSyncSetupScreen> {
  late final TextEditingController _organizationController;
  late final TextEditingController _computerController;
  late final TextEditingController _folderController;
  bool _saving = false;

  @override
  void initState() {
    super.initState();
    final state = context.read<AppState>();
    _organizationController = TextEditingController(
      text: state.workspaceProvider == 'desktop' ? state.workspaceName : '',
    );
    _computerController = TextEditingController(text: state.desktopComputerLabel);
    _folderController = TextEditingController(
      text: state.desktopFolderLabel.trim().isEmpty ? 'GrantProof' : state.desktopFolderLabel,
    );
  }

  @override
  void dispose() {
    _organizationController.dispose();
    _computerController.dispose();
    _folderController.dispose();
    super.dispose();
  }

  Future<void> _saveAndContinue() async {
    if (_saving) return;
    final s = AppStrings.of(context);
    final organizationName = _organizationController.text.trim().isEmpty
        ? 'GrantProof'
        : _organizationController.text.trim();
    final computerLabel = _computerController.text.trim();
    final folderLabel = _folderController.text.trim().isEmpty ? 'GrantProof' : _folderController.text.trim();

    setState(() => _saving = true);
    try {
      await context.read<AppState>().connectDesktopSync(
            organizationName: organizationName,
            computerLabel: computerLabel,
            folderLabel: folderLabel,
          );
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            s.isFr
                ? 'GrantProof Desktop Sync configuré pour $organizationName'
                : 'GrantProof Desktop Sync configured for $organizationName',
          ),
        ),
      );
      await Navigator.of(context).pushReplacement(
        MaterialPageRoute<void>(builder: (_) => const DesktopSyncPairingScreen()),
      );
    } finally {
      if (mounted) {
        setState(() => _saving = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);

    return Scaffold(
      appBar: AppBar(
        title: Text(s.isFr ? 'Configurer Desktop Sync' : 'Configure Desktop Sync'),
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
                        color: AppTheme.primary.withValues(alpha: 0.08),
                        borderRadius: BorderRadius.circular(16),
                      ),
                      child: const Icon(Icons.settings_rounded, color: AppTheme.primary),
                    ),
                    const SizedBox(width: 12),
                    Expanded(
                      child: Text(
                        s.isFr ? 'Étape 2 · Préparer ce téléphone' : 'Step 2 · Prepare this phone',
                        style: Theme.of(context).textTheme.titleMedium,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 12),
                Text(
                  s.isFr
                      ? 'Renseignez le nom de l’organisation, l’ordinateur cible et le dossier local. Juste après, vous passerez au jumelage avec le QR affiché par le compagnon PC.'
                      : 'Enter the organization name, the target computer, and the local folder. Right after that, you will move to pairing with the QR displayed by the desktop companion.',
                ),
              ],
            ),
          ),
          const SizedBox(height: 18),
          TextField(
            controller: _organizationController,
            decoration: InputDecoration(
              labelText: s.isFr ? 'Nom de l’organisation' : 'Organization name',
            ),
          ),
          const SizedBox(height: 12),
          TextField(
            controller: _computerController,
            decoration: InputDecoration(
              labelText: s.isFr ? 'Nom de l’ordinateur' : 'Computer label',
              hintText: s.isFr ? 'Ex. Bureau plaidoyer' : 'E.g. Advocacy desk',
            ),
          ),
          const SizedBox(height: 12),
          TextField(
            controller: _folderController,
            decoration: InputDecoration(
              labelText: s.isFr ? 'Nom du dossier local' : 'Local folder name',
              hintText: 'GrantProof',
            ),
          ),
          const SizedBox(height: 18),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  s.isFr ? 'Ce qui se passe ensuite' : 'What happens next',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 10),
                Text(
                  s.isFr
                      ? 'Une fois cette configuration enregistrée, ouvrez GrantProof Desktop Sync sur le PC, affichez son QR de jumelage et scannez-le depuis le téléphone.'
                      : 'Once this setup is saved, open GrantProof Desktop Sync on the PC, display its pairing QR, and scan it from the phone.',
                ),
              ],
            ),
          ),
          const SizedBox(height: 18),
          Row(
            children: [
              Expanded(
                child: OutlinedButton(
                  onPressed: _saving ? null : () => Navigator.of(context).pop(),
                  child: Text(s.cancel),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: FilledButton.icon(
                  onPressed: _saving ? null : _saveAndContinue,
                  icon: const Icon(Icons.arrow_forward_rounded),
                  label: Text(
                    _saving
                        ? (s.isFr ? 'Configuration…' : 'Configuring…')
                        : (s.isFr ? 'Continuer vers le jumelage' : 'Continue to pairing'),
                  ),
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}
