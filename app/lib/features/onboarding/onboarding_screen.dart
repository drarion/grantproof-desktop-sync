import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/localization/app_strings.dart';
import '../../core/state/app_state.dart';
import '../../shared/widgets/frosted_card.dart';

class OnboardingScreen extends StatefulWidget {
  const OnboardingScreen({super.key});

  @override
  State<OnboardingScreen> createState() => _OnboardingScreenState();
}

class _OnboardingScreenState extends State<OnboardingScreen> {
  final PageController _controller = PageController();
  int _index = 0;

  @override
  Widget build(BuildContext context) {
    final s = AppStrings.of(context);
    final slides = [
      (Icons.photo_camera_back_outlined, s.slide1Title, s.slide1Body),
      (Icons.cloud_off_outlined, s.slide2Title, s.slide2Body),
      (Icons.picture_as_pdf_outlined, s.slide3Title, s.slide3Body),
    ];

    return Scaffold(
      body: SafeArea(
        child: Padding(
          padding: const EdgeInsets.all(24),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Row(
                children: [
                  const Spacer(),
                  const _LanguageToggle(),
                  const SizedBox(width: 8),
                  TextButton(
                    onPressed: () => context.read<AppState>().completeOnboarding(),
                    child: Text(s.skip),
                  ),
                ],
              ),
              Expanded(
                child: PageView.builder(
                  controller: _controller,
                  itemCount: slides.length,
                  onPageChanged: (value) => setState(() => _index = value),
                  itemBuilder: (context, index) {
                    final slide = slides[index];
                    return Center(
                      child: FrostedCard(
                        padding: const EdgeInsets.all(28),
                        child: Column(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            Icon(slide.$1, size: 64),
                            const SizedBox(height: 24),
                            Text(slide.$2, textAlign: TextAlign.center, style: Theme.of(context).textTheme.headlineMedium),
                            const SizedBox(height: 12),
                            Text(slide.$3, textAlign: TextAlign.center, style: Theme.of(context).textTheme.bodyLarge),
                          ],
                        ),
                      ),
                    );
                  },
                ),
              ),
              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: List.generate(
                  slides.length,
                  (index) => Container(
                    margin: const EdgeInsets.symmetric(horizontal: 4),
                    width: _index == index ? 22 : 8,
                    height: 8,
                    decoration: BoxDecoration(
                      color: _index == index ? Theme.of(context).colorScheme.primary : Theme.of(context).dividerColor,
                      borderRadius: BorderRadius.circular(20),
                    ),
                  ),
                ),
              ),
              const SizedBox(height: 24),
              ElevatedButton(
                onPressed: () async {
                  if (_index == slides.length - 1) {
                    await context.read<AppState>().completeOnboarding();
                  } else {
                    await _controller.nextPage(duration: const Duration(milliseconds: 260), curve: Curves.easeOut);
                  }
                },
                child: Text(_index == slides.length - 1 ? s.enterApp : s.continueLabel),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _LanguageToggle extends StatelessWidget {
  const _LanguageToggle();

  @override
  Widget build(BuildContext context) {
    final state = context.watch<AppState>();
    return Container(
      padding: const EdgeInsets.all(4),
      decoration: BoxDecoration(
        color: Theme.of(context).colorScheme.surface,
        borderRadius: BorderRadius.circular(18),
        border: Border.all(color: Theme.of(context).dividerColor),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          _LangChip(
            label: 'FR',
            isSelected: state.languageCode == 'fr',
            onTap: () => state.setLanguage('fr'),
          ),
          _LangChip(
            label: 'ENG',
            isSelected: state.languageCode == 'en',
            onTap: () => state.setLanguage('en'),
          ),
        ],
      ),
    );
  }
}

class _LangChip extends StatelessWidget {
  final String label;
  final bool isSelected;
  final VoidCallback onTap;

  const _LangChip({required this.label, required this.isSelected, required this.onTap});

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onTap: onTap,
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
        decoration: BoxDecoration(
          color: isSelected ? Theme.of(context).colorScheme.primary : Colors.transparent,
          borderRadius: BorderRadius.circular(14),
        ),
        child: Text(
          label,
          style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                color: isSelected ? Colors.white : null,
                fontWeight: FontWeight.w700,
              ),
        ),
      ),
    );
  }
}
