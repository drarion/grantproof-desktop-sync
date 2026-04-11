import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../shared/widgets/frosted_card.dart';
import 'desktop_sync_setup_screen.dart';

class DesktopSyncInstallScreen extends StatelessWidget {
  const DesktopSyncInstallScreen({super.key});

  Future<void> _copyText(BuildContext context, String text, String successMessage) async {
    if (text.trim().isEmpty) return;
    await Clipboard.setData(ClipboardData(text: text));
    if (!context.mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(successMessage)));
  }

  Future<void> _copyInstallChecklist(BuildContext context, AppStrings s) async {
    final text = s.isFr
        ? 'GrantProof Desktop Sync\n\n1. Téléchargez ou récupérez GrantProofDesktopSync.exe sur le PC cible.\n2. Lancez le logiciel et choisissez le dossier local GrantProof.\n3. Ouvrez GrantProof Mobile, configurez Desktop Sync, puis scannez le QR affiché sur le PC.'
        : 'GrantProof Desktop Sync\n\n1. Download or retrieve GrantProofDesktopSync.exe on the target PC.\n2. Launch the app and choose the local GrantProof folder.\n3. Open GrantProof Mobile, configure Desktop Sync, then scan the QR shown on the PC.';
    await _copyText(
      context,
      text,
      s.isFr ? 'Checklist d’installation copiée.' : 'Installation checklist copied.',
    );
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final state = context.watch<AppState>();
    final downloadUrl = state.suggestedDesktopDownloadUrl;

    return Scaffold(
      appBar: AppBar(
        title: Text(s.isFr ? 'Installer le compagnon PC' : 'Install the desktop companion'),
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 32),
        children: [
          Container(
            decoration: BoxDecoration(
              borderRadius: BorderRadius.circular(30),
              gradient: const LinearGradient(
                begin: Alignment.topLeft,
                end: Alignment.bottomRight,
                colors: [Color(0xFF0B2D77), Color(0xFF1B4EB3), Color(0xFFFF7B1A)],
              ),
              boxShadow: const [
                BoxShadow(color: Color(0x220B2D77), blurRadius: 32, offset: Offset(0, 16)),
              ],
            ),
            padding: const EdgeInsets.all(22),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
                  decoration: BoxDecoration(
                    color: Colors.white.withValues(alpha: 0.14),
                    borderRadius: BorderRadius.circular(18),
                    border: Border.all(color: Colors.white.withValues(alpha: 0.14)),
                  ),
                  child: Text(
                    'GrantProof Desktop Sync',
                    style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                          color: Colors.white,
                          fontWeight: FontWeight.w700,
                        ),
                  ),
                ),
                const SizedBox(height: 14),
                Text(
                  s.isFr ? 'Le compagnon PC de GrantProof' : 'The PC companion for GrantProof',
                  style: Theme.of(context).textTheme.headlineMedium?.copyWith(color: Colors.white),
                ),
                const SizedBox(height: 10),
                Text(
                  s.isFr
                      ? 'Installez-le sur l’ordinateur du staff, affichez le QR de jumelage, puis laissez le téléphone synchroniser vers le dossier local GrantProof.'
                      : 'Install it on the staff computer, display the pairing QR, then let the phone sync into the local GrantProof folder.',
                  style: Theme.of(context).textTheme.bodyLarge?.copyWith(
                        color: Colors.white.withValues(alpha: 0.88),
                      ),
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
                  s.isFr ? 'Ce qu’il faut récupérer sur le PC' : 'What you need on the PC',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 10),
                Text(
                  s.isFr
                      ? 'Le logiciel à installer est GrantProofDesktopSync.exe. Il s’exécute sur Windows et prépare un dossier local GrantProof pour recevoir les preuves.'
                      : 'The software to install is GrantProofDesktopSync.exe. It runs on Windows and prepares a local GrantProof folder to receive the evidence.',
                ),
                if (downloadUrl.isNotEmpty) ...[
                  const SizedBox(height: 14),
                  SelectableText(
                    downloadUrl,
                    style: Theme.of(context).textTheme.bodySmall?.copyWith(color: AppTheme.primary),
                  ),
                  const SizedBox(height: 12),
                  OutlinedButton.icon(
                    onPressed: () => _copyText(
                      context,
                      downloadUrl,
                      s.isFr ? 'Lien d’installation copié.' : 'Install link copied.',
                    ),
                    icon: const Icon(Icons.copy_all_outlined),
                    label: Text(s.isFr ? 'Copier le lien' : 'Copy link'),
                  ),
                ] else ...[
                  const SizedBox(height: 14),
                  Text(
                    s.isFr
                        ? 'Aucun lien public n’est enregistré dans cette version. Utilisez l’installateur Windows déjà fourni par votre équipe ou votre dépôt GitHub Releases.'
                        : 'No public link is registered in this version. Use the Windows installer already shared by your team or your GitHub Releases page.',
                  ),
                ],
              ],
            ),
          ),
          const SizedBox(height: 18),
          FrostedCard(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  s.isFr ? 'Parcours recommandé' : 'Recommended flow',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 12),
                _GuideStep(
                  index: '1',
                  title: s.isFr ? 'Installer sur le PC cible' : 'Install on the target PC',
                  body: s.isFr
                      ? 'Lancez GrantProofDesktopSync.exe sur l’ordinateur du membre du staff et laissez-le créer ou pointer le dossier local GrantProof.'
                      : 'Launch GrantProofDesktopSync.exe on the staff computer and let it create or point to the local GrantProof folder.',
                ),
                const SizedBox(height: 12),
                _GuideStep(
                  index: '2',
                  title: s.isFr ? 'Configurer ce téléphone' : 'Configure this phone',
                  body: s.isFr
                      ? 'Dans GrantProof Mobile, renseignez le nom de l’organisation, le nom du PC et le nom du dossier local.'
                      : 'In GrantProof Mobile, enter the organization name, the PC name, and the local folder name.',
                ),
                const SizedBox(height: 12),
                _GuideStep(
                  index: '3',
                  title: s.isFr ? 'Scanner le QR du PC' : 'Scan the PC QR code',
                  body: s.isFr
                      ? 'Ouvrez l’écran de jumelage du compagnon PC puis scannez son QR depuis le téléphone pour activer la synchronisation locale.'
                      : 'Open the pairing screen in the desktop companion and scan its QR from the phone to activate local sync.',
                ),
              ],
            ),
          ),
          const SizedBox(height: 18),
          Wrap(
            spacing: 10,
            runSpacing: 10,
            children: [
              FilledButton.icon(
                onPressed: () => Navigator.of(context).push(
                  MaterialPageRoute<void>(builder: (_) => const DesktopSyncSetupScreen()),
                ),
                icon: const Icon(Icons.arrow_forward_rounded),
                label: Text(s.isFr ? 'Continuer vers la configuration' : 'Continue to setup'),
              ),
              OutlinedButton.icon(
                onPressed: () => _copyInstallChecklist(context, s),
                icon: const Icon(Icons.copy_all_outlined),
                label: Text(s.isFr ? 'Copier les étapes' : 'Copy the steps'),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _GuideStep extends StatelessWidget {
  final String index;
  final String title;
  final String body;

  const _GuideStep({
    required this.index,
    required this.title,
    required this.body,
  });

  @override
  Widget build(BuildContext context) {
    return Row(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Container(
          width: 30,
          height: 30,
          decoration: BoxDecoration(
            color: AppTheme.primary.withValues(alpha: 0.1),
            borderRadius: BorderRadius.circular(999),
          ),
          alignment: Alignment.center,
          child: Text(
            index,
            style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                  color: AppTheme.primary,
                  fontWeight: FontWeight.w800,
                ),
          ),
        ),
        const SizedBox(width: 12),
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(title, style: Theme.of(context).textTheme.titleMedium),
              const SizedBox(height: 4),
              Text(body),
            ],
          ),
        ),
      ],
    );
  }
}
