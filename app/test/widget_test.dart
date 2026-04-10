import 'package:flutter_test/flutter_test.dart';
import 'package:grantproof/main.dart';

void main() {
  testWidgets('app boots', (tester) async {
    await tester.pumpWidget(const GrantProofBootstrap());
    expect(find.text('GrantProof'), findsWidgets);
  });
}
