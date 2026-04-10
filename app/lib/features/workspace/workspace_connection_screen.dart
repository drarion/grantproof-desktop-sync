import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:provider/provider.dart';
import 'package:qr_flutter/qr_flutter.dart';

import '../../core/localization/app_strings.dart';
import '../../core/services/google_drive_workspace_service.dart';
import '../../core/services/microsoft365_workspace_service.dart';
import '../../core/services/desktop_sync_service.dart';
import '../../core/state/app_state.dart';
import '../../core/theme/app_theme.dart';
import '../../shared/widgets/frosted_card.dart';
import 'desktop_sync_install_screen.dart';
import 'desktop_sync_pairing_screen.dart';
import 'desktop_sync_setup_screen.dart';
import 'provisioning_scanner_screen.dart';

class WorkspaceConnectionScreen extends StatefulWidget {
  const WorkspaceConnectionScreen({super.key});

  @override
  State<WorkspaceConnectionScreen> createState() => _WorkspaceConnectionScreenState();
}

class _WorkspaceConnectionScreenState extends State<WorkspaceConnectionScreen> {
  bool _googleBusy = false;
  bool _microsoftBusy = false;

  Future<void> _copyText(String text, String successMessage) async {
    if (text.trim().isEmpty) return;
    await Clipboard.setData(ClipboardData(text: text));
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(successMessage)));
  }

  Future<void> _connectGoogleWorkspace({bool authorizeOnly = false}) async {
    final messenger = ScaffoldMessenger.of(context);
    final s = AppStrings.of(context);
    final state = context.read<AppState>();
    setState(() => _googleBusy = true);
    try {
      final account = await GoogleDriveWorkspaceService.instance.signInInteractive();
      if (account == null) {
        return;
      }

      if (authorizeOnly && state.workspaceDriveId.isNotEmpty) {
        await state.authorizeCurrentWorkspace(userEmail: account.email);
        if (!mounted) return;
        messenger.showSnackBar(
          SnackBar(
            content: Text(
              s.isFr
                  ? 'Compte Google connecté pour cet appareil : ${account.email}'
                  : 'Google account connected for this device: ${account.email}',
            ),
          ),
        );
        return;
      }

      final drives = await GoogleDriveWorkspaceService.instance.listSharedDrives();
      if (!mounted) return;
      if (drives.isEmpty) {
        messenger.showSnackBar(
          SnackBar(
            content: Text(
              s.isFr
                  ? 'Aucun Shared Drive visible pour ce compte. Vérifie les accès de l’utilisateur dans Google Workspace.'
                  : 'No Shared Drive is visible for this account. Check the user permissions in Google Workspace.',
            ),
          ),
        );
        return;
      }

      final selected = await showModalBottomSheet<GoogleSharedDriveInfo>(
        context: context,
        showDragHandle: true,
        builder: (context) => _GoogleDrivePicker(drives: drives),
      );
      if (!mounted || selected == null) return;

      await state.connectGoogleSharedDrive(
        driveId: selected.id,
        driveName: selected.name,
        userEmail: account.email,
      );
      await state.regenerateOrganizationProvisioningCode();

      if (!mounted) return;
      messenger.showSnackBar(
        SnackBar(
          content: Text(
            s.isFr
                ? 'Google Workspace configuré : ${selected.name}'
                : 'Google Workspace configured: ${selected.name}',
          ),
        ),
      );
    } catch (error) {
      if (!mounted) return;
      final rawError = error.toString();
      final isCode10 = rawError.contains('sign_in_failed') && rawError.contains(': 10');
      final message = isCode10
          ? (s.isFr
              ? 'Connexion Google impossible : code 10. Vérifie le package Android com.grantproof et le SHA-1 du certificat signé dans Google Cloud.'
              : 'Google connection failed: code 10. Check the Android package com.grantproof and the signed certificate SHA-1 in Google Cloud.')
          : (s.isFr ? 'Connexion Google impossible : $error' : 'Google connection failed: $error');
      messenger.showSnackBar(SnackBar(content: Text(message)));
    } finally {
      if (mounted) {
        setState(() => _googleBusy = false);
      }
    }
  }

  Future<void> _connectMicrosoftWorkspace({bool authorizeOnly = false}) async {
    final messenger = ScaffoldMessenger.of(context);
    final s = AppStrings.of(context);
    final state = context.read<AppState>();

    String tenantId = state.microsoftTenantId;
    String clientId = state.microsoftClientId;

    if (!authorizeOnly) {
      final config = await showModalBottomSheet<_MicrosoftAuthConfig>(
        context: context,
        isScrollControlled: true,
        showDragHandle: true,
        builder: (context) => _MicrosoftConfigSheet(
          initialTenantId: state.microsoftTenantId,
          initialClientId: state.microsoftClientId,
        ),
      );
      if (!mounted || config == null) return;
      if (config.clientId.trim().isEmpty) {
        messenger.showSnackBar(SnackBar(content: Text(s.microsoftNeedConfig)));
        return;
      }

      tenantId = config.tenantId;
      clientId = config.clientId;
      await state.saveMicrosoftAuthConfig(tenantId: tenantId, clientId: clientId);
    } else if (clientId.trim().isEmpty) {
      messenger.showSnackBar(SnackBar(content: Text(s.microsoftNeedConfig)));
      return;
    }

    setState(() => _microsoftBusy = true);
    try {
      final session = await Microsoft365WorkspaceService.instance.signInInteractive(
        tenantId: tenantId,
        clientId: clientId,
      );
      if (!mounted) return;

      if (authorizeOnly && state.workspaceSiteId.isNotEmpty && state.workspaceDriveId.isNotEmpty) {
        await state.authorizeCurrentWorkspace(userEmail: session.email);
        if (!mounted) return;
        messenger.showSnackBar(
          SnackBar(
            content: Text(
              s.isFr
                  ? 'Compte Microsoft connecté pour cet appareil : ${session.email}'
                  : 'Microsoft account connected for this device: ${session.email}',
            ),
          ),
        );
        return;
      }

      final site = await showModalBottomSheet<Microsoft365SiteInfo>(
        context: context,
        isScrollControlled: true,
        showDragHandle: true,
        builder: (context) => const _MicrosoftSitePickerSheet(),
      );
      if (!mounted || site == null) return;

      final drives = await Microsoft365WorkspaceService.instance.listSiteDrives(site.id);
      if (!mounted) return;
      if (drives.isEmpty) {
        messenger.showSnackBar(SnackBar(content: Text(s.microsoftNoLibrariesFound)));
        return;
      }

      final drive = await showModalBottomSheet<Microsoft365DriveInfo>(
        context: context,
        showDragHandle: true,
        builder: (context) => _MicrosoftDrivePicker(drives: drives),
      );
      if (!mounted || drive == null) return;

      await state.connectMicrosoftSharePoint(
        siteId: site.id,
        siteName: site.name,
        siteUrl: site.webUrl,
        driveId: drive.id,
        driveName: drive.name,
        userEmail: session.email,
      );
      await state.regenerateOrganizationProvisioningCode();

      if (!mounted) return;
      messenger.showSnackBar(
        SnackBar(
          content: Text(
            s.isFr
                ? 'Microsoft 365 configuré : ${site.name} / ${drive.name}'
                : 'Microsoft 365 configured: ${site.name} / ${drive.name}',
          ),
        ),
      );
    } catch (error) {
      if (!mounted) return;
      messenger.showSnackBar(
        SnackBar(
          content: Text(
            s.isFr ? 'Connexion Microsoft impossible : $error' : 'Microsoft connection failed: $error',
          ),
        ),
      );
    } finally {
      if (mounted) {
        setState(() => _microsoftBusy = false);
      }
    }
  }

  Future<void> _openDesktopInstallGuide() async {
    await Navigator.of(context).push(
      MaterialPageRoute<void>(builder: (_) => const DesktopSyncInstallScreen()),
    );
    if (mounted) {
      setState(() {});
    }
  }

  Future<void> _setupDesktopSync() async {
    await Navigator.of(context).push(
      MaterialPageRoute<void>(builder: (_) => const DesktopSyncSetupScreen()),
    );
    if (mounted) {
      setState(() {});
    }
  }

  Future<void> _openDesktopPairingScreen() async {
    await Navigator.of(context).push(
      MaterialPageRoute<void>(builder: (_) => const DesktopSyncPairingScreen()),
    );
    if (mounted) {
      setState(() {});
    }
  }

  Future<void> _handleProvisioningPayload(String rawPayload) async {
    final s = AppStrings.of(context);
    final state = context.read<AppState>();
    final messenger = ScaffoldMessenger.of(context);
    final payload = rawPayload.trim();
    if (payload.isEmpty) return;

    try {
      if (DesktopSyncService.instance.looksLikePairingPayload(payload)) {
        await state.pairDesktopCompanion(payload);
        if (!mounted) return;
        messenger.showSnackBar(
          SnackBar(
            content: Text(
              s.isFr
                  ? 'Ordinateur jumelé : ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}'
                  : 'Computer paired: ${state.desktopComputerLabel.isEmpty ? state.workspaceName : state.desktopComputerLabel}',
            ),
          ),
        );
        return;
      }

      await state.importOrganizationCode(payload);
      if (!mounted) return;
      messenger.showSnackBar(
        SnackBar(
          content: Text(
            s.isFr
                ? 'Configuration d’organisation importée sur cet appareil.'
                : 'Organization configuration imported on this device.',
          ),
        ),
      );
    } catch (error) {
      if (!mounted) return;
      messenger.showSnackBar(
        SnackBar(
          content: Text(
            s.isFr ? 'Code invalide : $error' : 'Invalid code: $error',
          ),
        ),
      );
    }
  }

  Future<void> _importOrganizationCode() async {
    final code = await showModalBottomSheet<String>(
      context: context,
      isScrollControlled: true,
      showDragHandle: true,
      builder: (context) => const _OrganizationCodeImportSheet(),
    );
    if (!mounted || code == null) return;
    await _handleProvisioningPayload(code);
  }

  Future<void> _scanProvisioningQr() async {
    final raw = await Navigator.of(context).push<String>(
      MaterialPageRoute<String>(
        builder: (_) => ProvisioningScannerScreen(
          title: AppStrings.of(context).isFr ? 'Scanner un QR GrantProof' : 'Scan a GrantProof QR',
        ),
      ),
    );
    if (!mounted || raw == null) return;
    await _handleProvisioningPayload(raw);
  }

  String _statusLabel(AppState state, AppStrings s) {
    switch (state.workspaceStatusLabel) {
      case 'desktop_configured':
        return s.isFr ? 'Desktop prêt' : 'Desktop ready';
      case 'account_pending':
        return s.isFr ? 'Compte à connecter' : 'Account pending';
      case 'ready':
        return s.isFr ? 'Prêt' : 'Ready';
      default:
        return s.workspaceNotConnected;
    }
  }

  String _statusDescription(AppState state, AppStrings s) {
    if (!state.workspaceConnected) {
      return s.isFr
          ? 'Flux recommandé : installez GrantProof Desktop Sync sur le PC, jumeler ce téléphone avec le QR du PC, puis synchronisez vers le dossier local GrantProof.'
          : 'Recommended flow: install GrantProof Desktop Sync on the computer, pair this phone with the PC QR code, then sync to the local GrantProof folder.';
    }
    if (state.workspaceProvider == 'desktop') {
      if (state.desktopPairingReady) {
        return s.isFr
            ? 'Ce téléphone est jumelé à ${state.desktopComputerLabel.isEmpty ? 'votre ordinateur' : state.desktopComputerLabel}. Les synchronisations locales partent vers le dossier ${state.desktopFolderLabel}.'
            : 'This phone is paired with ${state.desktopComputerLabel.isEmpty ? 'your computer' : state.desktopComputerLabel}. Local syncs now flow to the ${state.desktopFolderLabel} folder.';
      }
      return s.isFr
          ? 'Desktop Sync est configuré. Il reste à ouvrir le compagnon PC puis à scanner son QR de jumelage.'
          : 'Desktop Sync is configured. You only need to open the desktop companion and scan its pairing QR.';
    }
    if (state.workspaceNeedsAccountConnection) {
      return s.isFr
          ? 'La configuration d’organisation est importée. Il reste simplement à connecter le compte de cet appareil.'
          : 'The organization setup is imported. You only need to connect the account used on this device.';
    }
    return s.isFr
        ? 'La configuration de stockage et le compte local de cet appareil sont prêts.'
        : 'The storage configuration and the local account on this device are ready.';
  }

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    final s = AppStrings.of(context);
    final isDesktopSelected = state.workspaceProvider == 'desktop' && state.workspaceConnected;
    final isAdvancedWorkspace = (state.workspaceProvider == 'google' || state.workspaceProvider == 'microsoft') && state.workspaceConnected;

    return Scaffold(
      appBar: AppBar(title: Text(s.workspace)),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 40),
        children: [
          _HeroStatusCard(
            title: s.workspace,
            subtitle: _statusDescription(state, s),
            statusLabel: _statusLabel(state, s),
            providerLabel: state.workspaceConnected ? state.workspaceProviderLabel : null,
            organizationName: state.workspaceConnected ? state.workspaceName : null,
            libraryLabel: state.workspaceConnected ? state.workspaceLibrary : null,
            userEmail: state.workspaceUserEmail.isNotEmpty ? state.workspaceUserEmail : null,
            extraLine: state.workspaceProvider == 'desktop'
                ? (state.desktopPairingReady
                    ? (s.isFr
                        ? 'Jumelé à ${state.desktopComputerLabel} • ${state.desktopServerHost}:${state.desktopServerPort}'
                        : 'Paired with ${state.desktopComputerLabel} • ${state.desktopServerHost}:${state.desktopServerPort}')
                    : (state.desktopComputerLabel.isEmpty
                        ? (s.isFr ? 'Aucun ordinateur jumelé pour l’instant.' : 'No paired computer yet.')
                        : (s.isFr
                            ? 'Ordinateur cible : ${state.desktopComputerLabel} • jumelage à faire'
                            : 'Target computer: ${state.desktopComputerLabel} • pairing still required')))
                : (state.workspaceProvider == 'microsoft' && state.workspaceSiteUrl.isNotEmpty
                    ? state.workspaceSiteUrl
                    : null),
            onDisconnect: state.workspaceConnected
                ? () => context.read<AppState>().disconnectWorkspace()
                : null,
          ),
          const SizedBox(height: 18),
          _DesktopPrimaryCard(
            configured: isDesktopSelected,
            paired: state.desktopPairingReady,
            onOpenInstallGuide: _openDesktopInstallGuide,
            onOpenSetup: _setupDesktopSync,
            onOpenPairing: _openDesktopPairingScreen,
            organizationName: isDesktopSelected ? state.workspaceName : '',
            computerLabel: state.desktopComputerLabel,
            folderLabel: state.desktopFolderLabel,
          ),
          const SizedBox(height: 22),
          Text(
            s.isFr ? 'Options avancées' : 'Advanced options',
            style: Theme.of(context).textTheme.titleMedium,
          ),
          const SizedBox(height: 8),
          Text(
            s.isFr
                ? 'Google Workspace et Microsoft 365 restent disponibles pour les organisations qui veulent un espace partagé plus structuré.'
                : 'Google Workspace and Microsoft 365 remain available for organizations that want a more structured shared workspace.',
          ),
          const SizedBox(height: 14),
          Row(
            children: [
              Expanded(
                child: _ProviderMiniCard(
                  accent: AppTheme.primary,
                  icon: Icons.drive_folder_upload_outlined,
                  title: 'Google Workspace',
                  subtitle: state.workspaceProvider == 'google' && state.workspaceConnected
                      ? (state.workspaceNeedsAccountConnection
                          ? (s.isFr ? 'Configuration importée • compte local requis' : 'Imported setup • local account required')
                          : (s.isFr ? 'Shared Drive prêt pour cet appareil' : 'Shared Drive ready on this device'))
                      : (s.isFr ? 'Option avancée • Shared Drive' : 'Advanced option • Shared Drive'),
                  buttonLabel: state.workspaceProvider == 'google' && state.workspaceConnected && state.workspaceNeedsAccountConnection
                      ? (s.isFr ? 'Connecter mon compte' : 'Connect my account')
                      : (_googleBusy ? (s.isFr ? 'Connexion…' : 'Connecting…') : (s.isFr ? 'Configurer' : 'Configure')),
                  onTap: _googleBusy
                      ? null
                      : () => _connectGoogleWorkspace(
                            authorizeOnly: state.workspaceProvider == 'google' && state.workspaceConnected && state.workspaceNeedsAccountConnection,
                          ),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: _ProviderMiniCard(
                  accent: AppTheme.secondary,
                  icon: Icons.business_center_outlined,
                  title: 'Microsoft 365',
                  subtitle: state.workspaceProvider == 'microsoft' && state.workspaceConnected
                      ? (state.workspaceNeedsAccountConnection
                          ? (s.isFr ? 'Configuration importée • compte local requis' : 'Imported setup • local account required')
                          : (s.isFr ? 'Site et bibliothèque prêts sur cet appareil' : 'Site and library ready on this device'))
                      : (s.isFr ? 'Option entreprise • SharePoint' : 'Enterprise option • SharePoint'),
                  buttonLabel: state.workspaceProvider == 'microsoft' && state.workspaceConnected && state.workspaceNeedsAccountConnection
                      ? (s.isFr ? 'Connecter mon compte' : 'Connect my account')
                      : (_microsoftBusy ? (s.isFr ? 'Connexion…' : 'Connecting…') : (s.isFr ? 'Configurer' : 'Configure')),
                  onTap: _microsoftBusy
                      ? null
                      : () => _connectMicrosoftWorkspace(
                            authorizeOnly: state.workspaceProvider == 'microsoft' && state.workspaceConnected && state.workspaceNeedsAccountConnection,
                          ),
                ),
              ),
            ],
          ),
          if (isAdvancedWorkspace) ...[
            const SizedBox(height: 18),
            FrostedCard(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    children: [
                      Container(
                        width: 42,
                        height: 42,
                        decoration: BoxDecoration(
                          color: AppTheme.primary.withValues(alpha: 0.08),
                          borderRadius: BorderRadius.circular(16),
                        ),
                        child: const Icon(Icons.qr_code_2_rounded, color: AppTheme.primary),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: Text(
                          s.isFr ? 'Partager la configuration' : 'Share the configuration',
                          style: Theme.of(context).textTheme.titleMedium,
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  Text(
                    state.workspaceConnected
                        ? (s.isFr
                            ? 'Une fois la destination choisie, partagez ce code aux autres appareils de l’équipe. Ils importeront ensuite la configuration avant de connecter leur propre compte si nécessaire.'
                            : 'Once the destination is selected, share this code with the other devices on the team. They import the configuration and then connect their own account if needed.')
                        : (s.isFr
                            ? 'Configurez d’abord une destination, puis le code d’organisation sera généré ici.'
                            : 'Configure a destination first, then the organization code will appear here.'),
                  ),
                  if (state.organizationProvisioningCode.isNotEmpty) ...[
                    const SizedBox(height: 18),
                    Center(
                      child: Container(
                        padding: const EdgeInsets.all(12),
                        decoration: BoxDecoration(
                          color: Colors.white,
                          borderRadius: BorderRadius.circular(24),
                          border: Border.all(color: AppTheme.border),
                        ),
                        child: QrImageView(
                          data: state.organizationProvisioningCode,
                          version: QrVersions.auto,
                          size: 180,
                          eyeStyle: const QrEyeStyle(eyeShape: QrEyeShape.square, color: AppTheme.primary),
                          dataModuleStyle: const QrDataModuleStyle(dataModuleShape: QrDataModuleShape.square, color: AppTheme.text),
                        ),
                      ),
                    ),
                    const SizedBox(height: 14),
                    SelectableText(
                      state.organizationProvisioningCode,
                      style: Theme.of(context).textTheme.bodySmall,
                    ),
                    const SizedBox(height: 14),
                    Wrap(
                      spacing: 10,
                      runSpacing: 10,
                      children: [
                        FilledButton.tonalIcon(
                          onPressed: () => _copyText(
                            state.organizationProvisioningCode,
                            s.isFr ? 'Code d’organisation copié.' : 'Organization code copied.',
                          ),
                          icon: const Icon(Icons.copy_all_outlined),
                          label: Text(s.isFr ? 'Copier le code' : 'Copy code'),
                        ),
                        OutlinedButton.icon(
                          onPressed: () async {
                            final appState = context.read<AppState>();
                            final messenger = ScaffoldMessenger.of(context);
                            final code = await appState.regenerateOrganizationProvisioningCode();
                            if (!mounted) return;
                            messenger.showSnackBar(
                              SnackBar(
                                content: Text(
                                  s.isFr ? 'Code régénéré.' : 'Code regenerated.',
                                ),
                              ),
                            );
                            await _copyText(
                              code,
                              s.isFr ? 'Nouveau code copié.' : 'New code copied.',
                            );
                          },
                          icon: const Icon(Icons.refresh_outlined),
                          label: Text(s.isFr ? 'Régénérer' : 'Regenerate'),
                        ),
                      ],
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
                  Row(
                    children: [
                      Container(
                        width: 42,
                        height: 42,
                        decoration: BoxDecoration(
                          color: AppTheme.secondary.withValues(alpha: 0.1),
                          borderRadius: BorderRadius.circular(16),
                        ),
                        child: const Icon(Icons.phone_iphone_rounded, color: AppTheme.secondary),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: Text(
                          s.isFr ? 'Rejoindre une configuration' : 'Join a configuration',
                          style: Theme.of(context).textTheme.titleMedium,
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  Text(
                    s.isFr
                        ? 'Scannez ou collez le code partagé par votre équipe. Pour Google ou Microsoft, vous connecterez ensuite votre propre compte sur cet appareil.'
                        : 'Scan or paste the code shared by your team. For Google or Microsoft, you then connect your own account on this device.',
                  ),
                  const SizedBox(height: 16),
                  Wrap(
                    spacing: 10,
                    runSpacing: 10,
                    children: [
                      FilledButton.icon(
                        onPressed: _scanProvisioningQr,
                        icon: const Icon(Icons.qr_code_scanner_rounded),
                        label: Text(s.isFr ? 'Scanner un QR' : 'Scan a QR'),
                      ),
                      OutlinedButton.icon(
                        onPressed: _importOrganizationCode,
                        icon: const Icon(Icons.download_done_outlined),
                        label: Text(s.isFr ? 'Importer un code' : 'Import a code'),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          ],
        ],
      ),
    );
  }
}

class _HeroStatusCard extends StatelessWidget {
  final String title;
  final String subtitle;
  final String statusLabel;
  final String? providerLabel;
  final String? organizationName;
  final String? libraryLabel;
  final String? userEmail;
  final String? extraLine;
  final VoidCallback? onDisconnect;

  const _HeroStatusCard({
    required this.title,
    required this.subtitle,
    required this.statusLabel,
    this.providerLabel,
    this.organizationName,
    this.libraryLabel,
    this.userEmail,
    this.extraLine,
    this.onDisconnect,
  });

  @override
  Widget build(BuildContext context) {
    return Container(
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
          Text(
            title,
            style: Theme.of(context).textTheme.headlineMedium?.copyWith(color: Colors.white),
          ),
          const SizedBox(height: 10),
          Text(
            subtitle,
            style: Theme.of(context).textTheme.bodyLarge?.copyWith(color: Colors.white.withValues(alpha: 0.88)),
          ),
          const SizedBox(height: 10),
          Text(
            statusLabel,
            style: Theme.of(context).textTheme.titleSmall?.copyWith(
                  color: Colors.white,
                  fontWeight: FontWeight.w800,
                ),
          ),
          if (organizationName != null && organizationName!.trim().isNotEmpty) ...[
            const SizedBox(height: 18),
            Wrap(
              spacing: 10,
              runSpacing: 10,
              children: [
                if (providerLabel != null && providerLabel!.trim().isNotEmpty) _HeroPill(label: providerLabel!),
                if (organizationName != null && organizationName!.trim().isNotEmpty) _HeroPill(label: organizationName!),
                if (libraryLabel != null && libraryLabel!.trim().isNotEmpty) _HeroPill(label: libraryLabel!),
              ],
            ),
          ],
          if (userEmail != null || extraLine != null || onDisconnect != null) ...[
            const SizedBox(height: 16),
            if (userEmail != null && userEmail!.trim().isNotEmpty)
              Text(userEmail!, style: Theme.of(context).textTheme.bodyMedium?.copyWith(color: Colors.white70)),
            if (extraLine != null && extraLine!.trim().isNotEmpty) ...[
              const SizedBox(height: 6),
              Text(extraLine!, style: Theme.of(context).textTheme.bodyMedium?.copyWith(color: Colors.white70)),
            ],
            if (onDisconnect != null) ...[
              const SizedBox(height: 16),
              OutlinedButton.icon(
                onPressed: onDisconnect,
                icon: const Icon(Icons.link_off_outlined),
                label: Text(AppStrings.of(context).disconnectWorkspace),
                style: OutlinedButton.styleFrom(
                  foregroundColor: Colors.white,
                  side: BorderSide(color: Colors.white.withValues(alpha: 0.22)),
                ),
              ),
            ],
          ],
        ],
      ),
    );
  }
}

class _HeroPill extends StatelessWidget {
  final String label;

  const _HeroPill({required this.label});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
      decoration: BoxDecoration(
        color: Colors.white.withValues(alpha: 0.14),
        borderRadius: BorderRadius.circular(18),
        border: Border.all(color: Colors.white.withValues(alpha: 0.14)),
      ),
      child: Text(
        label,
        style: Theme.of(context).textTheme.bodyMedium?.copyWith(color: Colors.white, fontWeight: FontWeight.w700),
      ),
    );
  }
}

class _DesktopPrimaryCard extends StatelessWidget {
  final bool configured;
  final bool paired;
  final VoidCallback onOpenInstallGuide;
  final VoidCallback onOpenSetup;
  final VoidCallback onOpenPairing;
  final String organizationName;
  final String computerLabel;
  final String folderLabel;

  const _DesktopPrimaryCard({
    required this.configured,
    required this.paired,
    required this.onOpenInstallGuide,
    required this.onOpenSetup,
    required this.onOpenPairing,
    required this.organizationName,
    required this.computerLabel,
    required this.folderLabel,
  });

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final detailsStyle = Theme.of(context).textTheme.bodyMedium?.copyWith(
          color: AppTheme.text,
          fontWeight: FontWeight.w700,
        );

    return Container(
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(30),
        color: Colors.white,
        border: Border.all(color: AppTheme.border),
        boxShadow: const [
          BoxShadow(color: Color(0x120B2D77), blurRadius: 26, offset: Offset(0, 12)),
        ],
      ),
      padding: const EdgeInsets.all(20),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                width: 52,
                height: 52,
                decoration: BoxDecoration(
                  borderRadius: BorderRadius.circular(18),
                  gradient: const LinearGradient(
                    colors: [Color(0xFF0B2D77), Color(0xFFFF7B1A)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  ),
                ),
                child: const Icon(Icons.computer_rounded, color: Colors.white),
              ),
              const SizedBox(width: 14),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      'GrantProof Desktop Sync',
                      style: Theme.of(context).textTheme.titleLarge,
                    ),
                    const SizedBox(height: 4),
                    Text(
                      s.isFr
                          ? 'Le flux principal : installer le compagnon PC, jumeler le téléphone avec le QR du PC, puis synchroniser vers le dossier local GrantProof.'
                          : 'The main flow: install the desktop companion, pair the phone with the PC QR code, then sync to the local GrantProof folder.',
                    ),
                  ],
                ),
              ),
            ],
          ),
          const SizedBox(height: 16),
          Wrap(
            spacing: 10,
            runSpacing: 10,
            children: [
              _StepPill(label: s.isFr ? '1. Installer le compagnon PC' : '1. Install the desktop companion'),
              _StepPill(label: s.isFr ? '2. Jumeler le téléphone' : '2. Pair the phone'),
              _StepPill(label: s.isFr ? '3. Synchroniser vers GrantProof' : '3. Sync to GrantProof'),
            ],
          ),
          if (configured) ...[
            const SizedBox(height: 16),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                color: AppTheme.surfaceMuted,
                borderRadius: BorderRadius.circular(22),
                border: Border.all(color: AppTheme.border),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    s.isFr ? 'Configuration actuelle' : 'Current setup',
                    style: Theme.of(context).textTheme.titleMedium,
                  ),
                  const SizedBox(height: 8),
                  Text(
                    s.isFr ? 'Organisation : $organizationName' : 'Organization: $organizationName',
                    style: detailsStyle,
                  ),
                  const SizedBox(height: 4),
                  Text(
                    s.isFr
                        ? 'Ordinateur : ${computerLabel.isEmpty ? 'À renseigner' : computerLabel}'
                        : 'Computer: ${computerLabel.isEmpty ? 'To be set' : computerLabel}',
                    style: detailsStyle,
                  ),
                  const SizedBox(height: 4),
                  Text(
                    s.isFr ? 'Dossier : $folderLabel' : 'Folder: $folderLabel',
                    style: detailsStyle,
                  ),
                  const SizedBox(height: 4),
                  Text(
                    paired
                        ? (s.isFr ? 'Jumelage du téléphone : actif' : 'Phone pairing: active')
                        : (s.isFr ? 'Jumelage du téléphone : à faire' : 'Phone pairing: pending'),
                    style: detailsStyle,
                  ),
                ],
              ),
            ),
          ],
          const SizedBox(height: 16),
          Wrap(
            spacing: 10,
            runSpacing: 10,
            children: [
              FilledButton.icon(
                onPressed: onOpenInstallGuide,
                icon: const Icon(Icons.download_rounded),
                label: Text(s.isFr ? 'Installer le compagnon PC' : 'Install the desktop companion'),
              ),
              FilledButton.tonalIcon(
                onPressed: configured ? onOpenPairing : onOpenSetup,
                icon: Icon(configured ? Icons.qr_code_scanner_rounded : Icons.tune_rounded),
                label: Text(
                  configured
                      ? (paired
                          ? (s.isFr ? 'Voir le jumelage' : 'View pairing')
                          : (s.isFr ? 'Jumeler le téléphone' : 'Pair the phone'))
                      : (s.isFr ? 'Configurer Desktop Sync' : 'Configure Desktop Sync'),
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _StepPill extends StatelessWidget {
  final String label;

  const _StepPill({required this.label});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(16),
        color: AppTheme.surfaceMuted,
        border: Border.all(color: AppTheme.border),
      ),
      child: Text(label, style: Theme.of(context).textTheme.bodyMedium?.copyWith(color: AppTheme.text)),
    );
  }
}

class _ProviderMiniCard extends StatelessWidget {
  final Color accent;
  final IconData icon;
  final String title;
  final String subtitle;
  final String buttonLabel;
  final VoidCallback? onTap;

  const _ProviderMiniCard({
    required this.accent,
    required this.icon,
    required this.title,
    required this.subtitle,
    required this.buttonLabel,
    required this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    return FrostedCard(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(
            width: 44,
            height: 44,
            decoration: BoxDecoration(
              color: accent.withValues(alpha: 0.08),
              borderRadius: BorderRadius.circular(16),
            ),
            child: Icon(icon, color: accent),
          ),
          const SizedBox(height: 12),
          Text(title, style: Theme.of(context).textTheme.titleMedium),
          const SizedBox(height: 8),
          Text(subtitle),
          const SizedBox(height: 16),
          SizedBox(
            width: double.infinity,
            child: FilledButton.tonal(
              onPressed: onTap,
              child: FittedBox(
                fit: BoxFit.scaleDown,
                child: Text(
                  buttonLabel,
                  maxLines: 1,
                  softWrap: false,
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }
}

class _GoogleDrivePicker extends StatelessWidget {
  final List<GoogleSharedDriveInfo> drives;

  const _GoogleDrivePicker({required this.drives});

  @override
  Widget build(BuildContext context) {
    return SafeArea(
      child: ListView.separated(
        shrinkWrap: true,
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 24),
        itemBuilder: (context, index) {
          final drive = drives[index];
          return ListTile(
            contentPadding: EdgeInsets.zero,
            leading: const Icon(Icons.drive_folder_upload_outlined),
            title: Text(drive.name),
            subtitle: Text(drive.id),
            trailing: const Icon(Icons.chevron_right),
            onTap: () => Navigator.of(context).pop(drive),
          );
        },
        separatorBuilder: (_, __) => const Divider(height: 1),
        itemCount: drives.length,
      ),
    );
  }
}

class _MicrosoftAuthConfig {
  final String tenantId;
  final String clientId;

  const _MicrosoftAuthConfig({required this.tenantId, required this.clientId});
}

class _OrganizationCodeImportSheet extends StatefulWidget {
  const _OrganizationCodeImportSheet();

  @override
  State<_OrganizationCodeImportSheet> createState() => _OrganizationCodeImportSheetState();
}

class _OrganizationCodeImportSheetState extends State<_OrganizationCodeImportSheet> {
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
            s.isFr ? 'Importer un code GrantProof' : 'Import a GrantProof code',
            style: Theme.of(context).textTheme.titleLarge,
          ),
          const SizedBox(height: 8),
          Text(
            s.isFr
                ? 'Collez ici soit le code d’organisation partagé par votre équipe, soit le code de jumelage affiché par GrantProof Desktop Sync.'
                : 'Paste here either the organization code shared by your team or the pairing code displayed by GrantProof Desktop Sync.',
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _codeController,
            minLines: 4,
            maxLines: 8,
            decoration: InputDecoration(
              labelText: s.isFr ? 'Code de configuration' : 'Configuration code',
            ),
          ),
          const SizedBox(height: 16),
          Row(
            children: [
              Expanded(
                child: OutlinedButton(
                  onPressed: () => Navigator.of(context).pop(),
                  child: Text(s.cancel),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: FilledButton(
                  onPressed: () => Navigator.of(context).pop(_codeController.text.trim()),
                  child: Text(s.isFr ? 'Importer' : 'Import'),
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _MicrosoftConfigSheet extends StatefulWidget {
  final String initialTenantId;
  final String initialClientId;

  const _MicrosoftConfigSheet({required this.initialTenantId, required this.initialClientId});

  @override
  State<_MicrosoftConfigSheet> createState() => _MicrosoftConfigSheetState();
}

class _MicrosoftConfigSheetState extends State<_MicrosoftConfigSheet> {
  late final TextEditingController _tenantController;
  late final TextEditingController _clientController;

  @override
  void initState() {
    super.initState();
    _tenantController = TextEditingController(text: widget.initialTenantId);
    _clientController = TextEditingController(text: widget.initialClientId);
  }

  @override
  void dispose() {
    _tenantController.dispose();
    _clientController.dispose();
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
          Text(s.microsoftSetupTitle, style: Theme.of(context).textTheme.titleLarge),
          const SizedBox(height: 8),
          Text(s.microsoftConfigHint),
          const SizedBox(height: 16),
          TextField(
            controller: _tenantController,
            decoration: InputDecoration(
              labelText: s.microsoftTenantId,
              hintText: s.isFr ? 'Ex. NrDStudio.onmicrosoft.com' : 'E.g. contoso.onmicrosoft.com',
            ),
          ),
          const SizedBox(height: 12),
          TextField(
            controller: _clientController,
            decoration: InputDecoration(
              labelText: s.microsoftClientId,
              hintText: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
            ),
          ),
          const SizedBox(height: 16),
          Text(
            Microsoft365WorkspaceService.redirectUri,
            style: Theme.of(context).textTheme.bodySmall,
          ),
          const SizedBox(height: 16),
          Row(
            children: [
              Expanded(
                child: OutlinedButton(
                  onPressed: () => Navigator.of(context).pop(),
                  child: Text(s.cancel),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: FilledButton(
                  onPressed: () {
                    Navigator.of(context).pop(
                      _MicrosoftAuthConfig(
                        tenantId: _tenantController.text.trim(),
                        clientId: _clientController.text.trim(),
                      ),
                    );
                  },
                  child: Text(s.continueLabel),
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _MicrosoftSitePickerSheet extends StatefulWidget {
  const _MicrosoftSitePickerSheet();

  @override
  State<_MicrosoftSitePickerSheet> createState() => _MicrosoftSitePickerSheetState();
}

class _MicrosoftSitePickerSheetState extends State<_MicrosoftSitePickerSheet> {
  final TextEditingController _queryController = TextEditingController();
  final TextEditingController _siteUrlController = TextEditingController();
  bool _loading = false;
  String? _error;
  List<Microsoft365SiteInfo> _sites = const <Microsoft365SiteInfo>[];

  @override
  void dispose() {
    _queryController.dispose();
    _siteUrlController.dispose();
    super.dispose();
  }

  Future<void> _search() async {
    final s = AppStrings.of(context);
    final query = _queryController.text.trim();
    if (query.isEmpty) {
      setState(() {
        _error = s.isFr ? 'Saisis un mot-clé de site avant de lancer la recherche.' : 'Enter a site keyword before starting the search.';
        _sites = const <Microsoft365SiteInfo>[];
      });
      return;
    }
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final sites = await Microsoft365WorkspaceService.instance.searchSites(query: query);
      if (!mounted) return;
      setState(() {
        _sites = sites;
        _error = sites.isEmpty ? s.microsoftNoSitesFound : null;
      });
    } catch (error) {
      if (!mounted) return;
      setState(() {
        _sites = const <Microsoft365SiteInfo>[];
        _error = error.toString();
      });
    } finally {
      if (mounted) {
        setState(() => _loading = false);
      }
    }
  }

  Future<void> _useExactUrl() async {
    final s = AppStrings.of(context);
    final siteUrl = _siteUrlController.text.trim();
    if (siteUrl.isEmpty) {
      setState(() {
        _error = s.isFr
            ? "Colle l'URL complète du site SharePoint si la recherche ne trouve rien."
            : 'Paste the full SharePoint site URL if search finds nothing.';
      });
      return;
    }
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final site = await Microsoft365WorkspaceService.instance.resolveSiteByUrl(siteUrl);
      if (!mounted) return;
      Navigator.of(context).pop(site);
    } catch (error) {
      if (!mounted) return;
      setState(() {
        _error = error.toString();
      });
    } finally {
      if (mounted) {
        setState(() => _loading = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final bottom = MediaQuery.of(context).viewInsets.bottom;
    return SafeArea(
      child: Padding(
        padding: EdgeInsets.fromLTRB(20, 8, 20, bottom + 24),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(s.microsoftSearchSite, style: Theme.of(context).textTheme.titleLarge),
            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: TextField(
                    controller: _queryController,
                    decoration: InputDecoration(
                      labelText: s.microsoftSiteKeyword,
                      hintText: s.isFr ? 'Ex. programme, pays, région…' : 'E.g. program, country, region…',
                    ),
                    onSubmitted: (_) => _search(),
                  ),
                ),
                const SizedBox(width: 12),
                FilledButton(
                  onPressed: _loading ? null : _search,
                  child: Text(s.microsoftSearch),
                ),
              ],
            ),
            const SizedBox(height: 10),
            Text(
              s.isFr
                  ? "Si la recherche ne trouve rien, colle l'URL complète du site SharePoint."
                  : 'If search finds nothing, paste the full SharePoint site URL.',
              style: Theme.of(context).textTheme.bodySmall,
            ),
            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: TextField(
                    controller: _siteUrlController,
                    decoration: InputDecoration(
                      labelText: s.isFr ? 'URL du site SharePoint' : 'SharePoint site URL',
                      hintText: 'https://nrdstudio.sharepoint.com/sites/GrantProofTest',
                    ),
                    onSubmitted: (_) => _useExactUrl(),
                  ),
                ),
                const SizedBox(width: 12),
                OutlinedButton(
                  onPressed: _loading ? null : _useExactUrl,
                  child: Text(s.isFr ? 'Utiliser URL' : 'Use URL'),
                ),
              ],
            ),
            const SizedBox(height: 16),
            if (_loading)
              const Padding(
                padding: EdgeInsets.symmetric(vertical: 24),
                child: Center(child: CircularProgressIndicator()),
              )
            else if (_error != null)
              Padding(
                padding: const EdgeInsets.symmetric(vertical: 12),
                child: Text(_error!),
              )
            else if (_sites.isEmpty)
              Padding(
                padding: const EdgeInsets.symmetric(vertical: 12),
                child: Text(s.isFr ? 'Lance une recherche pour afficher les sites visibles.' : 'Run a search to display visible sites.'),
              )
            else
              SizedBox(
                height: 360,
                child: ListView.separated(
                  shrinkWrap: true,
                  itemCount: _sites.length,
                  separatorBuilder: (_, __) => const Divider(height: 1),
                  itemBuilder: (context, index) {
                    final site = _sites[index];
                    return ListTile(
                      contentPadding: EdgeInsets.zero,
                      leading: const Icon(Icons.business_outlined),
                      title: Text(site.name),
                      subtitle: Text(site.webUrl.isNotEmpty ? site.webUrl : site.id),
                      trailing: const Icon(Icons.chevron_right),
                      onTap: () => Navigator.of(context).pop(site),
                    );
                  },
                ),
              ),
          ],
        ),
      ),
    );
  }
}

class _MicrosoftDrivePicker extends StatelessWidget {
  final List<Microsoft365DriveInfo> drives;

  const _MicrosoftDrivePicker({required this.drives});

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    return SafeArea(
      child: Padding(
        padding: const EdgeInsets.fromLTRB(20, 8, 20, 24),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(s.microsoftChooseLibrary, style: Theme.of(context).textTheme.titleLarge),
            const SizedBox(height: 12),
            SizedBox(
              height: 360,
              child: ListView.separated(
                shrinkWrap: true,
                itemCount: drives.length,
                separatorBuilder: (_, __) => const Divider(height: 1),
                itemBuilder: (context, index) {
                  final drive = drives[index];
                  return ListTile(
                    contentPadding: EdgeInsets.zero,
                    leading: const Icon(Icons.folder_copy_outlined),
                    title: Text(drive.name),
                    subtitle: Text(drive.webUrl.isNotEmpty ? drive.webUrl : drive.id),
                    trailing: const Icon(Icons.chevron_right),
                    onTap: () => Navigator.of(context).pop(drive),
                  );
                },
              ),
            ),
          ],
        ),
      ),
    );
  }
}
