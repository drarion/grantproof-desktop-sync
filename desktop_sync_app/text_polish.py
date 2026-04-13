from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Iterable


COMMON_TYPO_REPLACEMENTS = {
    'fr': {
        'formasion': 'formation',
        'formaton': 'formation',
        'formatoin': 'formation',
        'formtion': 'formation',
        'beneficiare': 'bénéficiaire',
        'beneficiaires': 'bénéficiaires',
        'beneficaires': 'bénéficiaires',
        'sensibilisaton': 'sensibilisation',
        'sensibilisasion': 'sensibilisation',
        'organisaton': 'organisation',
        'organisasion': 'organisation',
        'reuinion': 'réunion',
        'reuinion': 'réunion',
        'communautairee': 'communautaire',
        'alphetisation': 'alphabétisation',
        'alphabetisasion': 'alphabétisation',
        'alphabetisation': 'alphabétisation',
        'temoignage': 'témoignage',
        'temoignages': 'témoignages',
        'equipemment': 'équipement',
        'equipements': 'équipements',
        'fomation': 'formation',
        'partcipants': 'participants',
        'atellier': 'atelier',
        'ateliere': 'atelier',
    },
    'en': {
        'trainning': 'training',
        'benificiary': 'beneficiary',
        'benificiaries': 'beneficiaries',
        'organizaton': 'organization',
        'organisaton': 'organization',
        'activty': 'activity',
        'commmunity': 'community',
        'metting': 'meeting',
        'equipement': 'equipment',
    },
}

LOWERCASE_EXCEPTIONS = {
    'fr': {'de', 'du', 'des', 'et', 'à', 'au', 'aux', 'en', 'pour', 'sur', 'dans', 'par', 'avec', 'ou', 'la', 'le', 'les', 'd’', "d'", 'l’', "l'"},
    'en': {'and', 'or', 'the', 'of', 'to', 'for', 'in', 'on', 'with', 'a', 'an'},
}

PROTECTED_ACRONYMS = {'ngo', 'eu', 'afd', 'usaid', 'wfp', 'undp', 'who', 'gps', 'm&e'}


@dataclass
class PolishedText:
    value: str
    original: str


class GrantProofTextPolisher:
    def normalize_label(self, value: object, *, lang: str = 'fr', choices: Iterable[str] | None = None, title_mode: bool = False) -> str:
        text = self._clean_text(value, lang=lang, title_mode=title_mode)
        matched = self._match_choice(text, choices or [])
        return matched or text

    def normalize_sentence(self, value: object, *, lang: str = 'fr') -> str:
        return self._clean_text(value, lang=lang, title_mode=False, sentence_mode=True)

    def normalize_quote(self, value: object, *, lang: str = 'fr') -> str:
        cleaned = self.normalize_sentence(value, lang=lang)
        return cleaned.rstrip('.') if cleaned else ''

    def normalize_country(self, value: object, *, lang: str = 'fr') -> str:
        return self._clean_text(value, lang=lang, title_mode=True)

    def _clean_text(self, value: object, *, lang: str, title_mode: bool = False, sentence_mode: bool = False) -> str:
        text = str(value or '').strip()
        if not text:
            return ''
        text = self._normalize_spacing(text)
        text = self._replace_common_typos(text, lang=lang)
        if title_mode:
            text = self._smart_title(text, lang=lang)
        elif sentence_mode:
            text = self._sentence_case(text)
        else:
            text = self._sentence_case(text)
        return text

    def _normalize_spacing(self, text: str) -> str:
        text = text.replace('\n', ' ')
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([,.;:!?])', r'\1', text)
        return text.strip()

    def _replace_common_typos(self, text: str, *, lang: str) -> str:
        replacements = COMMON_TYPO_REPLACEMENTS.get(lang, {})
        for wrong, right in replacements.items():
            pattern = re.compile(rf'(?i)\b{re.escape(wrong)}\b')
            text = pattern.sub(lambda m: self._preserve_case(m.group(0), right), text)
        return text

    def _preserve_case(self, source: str, replacement: str) -> str:
        if source.isupper():
            return replacement.upper()
        if source[:1].isupper():
            return replacement[:1].upper() + replacement[1:]
        return replacement

    def _sentence_case(self, text: str) -> str:
        segments = re.split(r'([.!?]\s+)', text)
        rebuilt: list[str] = []
        for index in range(0, len(segments), 2):
            sentence = segments[index].strip()
            sep = segments[index + 1] if index + 1 < len(segments) else ''
            if sentence:
                sentence = sentence[:1].upper() + sentence[1:]
            rebuilt.append(sentence + sep)
        return ''.join(rebuilt).strip()

    def _smart_title(self, text: str, *, lang: str) -> str:
        tokens = text.split(' ')
        rebuilt: list[str] = []
        for index, token in enumerate(tokens):
            bare = token.strip()
            if not bare:
                continue
            lower = bare.lower()
            if lower in PROTECTED_ACRONYMS:
                rebuilt.append(lower.upper())
                continue
            if lang == 'fr':
                if index == 0:
                    rebuilt.append(lower[:1].upper() + lower[1:])
                else:
                    rebuilt.append(lower)
                continue
            if index > 0 and lower in LOWERCASE_EXCEPTIONS.get(lang, set()):
                rebuilt.append(lower)
                continue
            rebuilt.append(lower[:1].upper() + lower[1:])
        return ' '.join(rebuilt)

    def _match_choice(self, value: str, choices: Iterable[str]) -> str | None:
        cleaned = value.strip()
        if not cleaned:
            return None
        normalized_value = self._normalize_key(cleaned)
        best_choice = None
        best_score = 0.0
        for choice in choices:
            candidate = str(choice or '').strip()
            if not candidate:
                continue
            score = SequenceMatcher(None, normalized_value, self._normalize_key(candidate)).ratio()
            if score > best_score:
                best_score = score
                best_choice = candidate
        if best_choice and best_score >= 0.72:
            return best_choice[:1].upper() + best_choice[1:]
        return None

    def _normalize_key(self, text: str) -> str:
        normalized = unicodedata.normalize('NFKD', text)
        ascii_text = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
        ascii_text = ascii_text.lower()
        ascii_text = re.sub(r'[^a-z0-9]+', ' ', ascii_text)
        return ascii_text.strip()
