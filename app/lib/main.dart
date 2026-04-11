import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import 'core/state/app_state.dart';
import 'core/theme/app_theme.dart';
import 'features/app_shell/app_shell.dart';
import 'features/onboarding/onboarding_screen.dart';
import 'features/splash/splash_screen.dart';

void main() {
  runApp(const GrantProofBootstrap());
}

class GrantProofBootstrap extends StatelessWidget {
  const GrantProofBootstrap({super.key});

  @override
  Widget build(BuildContext context) {
    return ChangeNotifierProvider(
      create: (_) => AppState()..bootstrap(),
      child: const GrantProofApp(),
    );
  }
}

class GrantProofApp extends StatelessWidget {
  const GrantProofApp({super.key});

  @override
  Widget build(BuildContext context) {
    return Consumer<AppState>(
      builder: (context, state, _) {
        return MaterialApp(
          title: 'GrantProof',
          debugShowCheckedModeBanner: false,
          theme: AppTheme.lightTheme,
          home: !state.isBootstrapped
              ? const SplashScreen()
              : state.onboardingCompleted
                  ? const AppShell()
                  : const OnboardingScreen(),
        );
      },
    );
  }
}
