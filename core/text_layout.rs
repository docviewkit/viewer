//! Format-independent horizontal text fitting. Adapters supply measured advances.

#[cfg(any(
    test,
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
pub(crate) fn punctuation_compression(character: char, advance: f32) -> f32 {
    if is_closing_punctuation(character)
        || matches!(
            character,
            '（' | '［' | '【' | '《' | '〈' | '「' | '『' | '〔' | '“' | '‘'
        )
    {
        advance.max(0.0) * 0.5
    } else {
        0.0
    }
}

/// Fit the boundary glyph using available punctuation whitespace. Once past the
/// natural edge, only closing punctuation can follow; do not greedily fill the
/// entire line at maximum compression.
#[cfg(any(
    test,
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
pub(crate) fn fits_compressed(
    used: f32,
    advance: f32,
    width: f32,
    capacity: f32,
    closing: bool,
) -> bool {
    (used <= width + 0.001 || closing) && used + advance - capacity <= width + 0.001
}

pub(crate) fn is_closing_punctuation(character: char) -> bool {
    matches!(
        character,
        '、' | '。'
            | '，'
            | '．'
            | '！'
            | '？'
            | '：'
            | '；'
            | '）'
            | '］'
            | '】'
            | '》'
            | '〉'
            | '」'
            | '』'
            | '〕'
            | '”'
            | '’'
    )
}

pub(crate) fn is_line_start_prohibited(character: char) -> bool {
    is_closing_punctuation(character)
        || matches!(
            character,
            ',' | '.' | ';' | ':' | '!' | '?' | ')' | ']' | '}' | '…'
        )
}

/// Reserve the glyph's ink, but allow up to half an em of punctuation whitespace
/// to be removed when a line would otherwise overflow. Never shrink all glyphs.
#[cfg(any(test, feature = "native-formats", feature = "legacy-office-formats"))]
pub(crate) fn compressed_line_breaks(
    characters: &[(char, f32)],
    first_width: f32,
    following_width: f32,
    default_tab_stop: f32,
) -> Vec<usize> {
    horizontal_line_breaks(
        characters,
        first_width,
        following_width,
        default_tab_stop,
        true,
    )
}

#[cfg(any(
    test,
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
pub(crate) fn horizontal_line_breaks(
    characters: &[(char, f32)],
    first_width: f32,
    following_width: f32,
    default_tab_stop: f32,
    compress_punctuation: bool,
) -> Vec<usize> {
    let mut breaks = Vec::new();
    let mut start = 0;
    let mut width = first_width.max(1.0);
    while start < characters.len() {
        let mut end = start;
        let mut used = 0.0_f32;
        let mut compression = 0.0;
        let mut has_tab = false;
        let mut whitespace = None;
        while end < characters.len() {
            let (character, mut advance) = characters[end];
            if character == '\n' {
                end += 1;
                break;
            }
            if character == '\t' {
                advance = default_tab_stop - used.rem_euclid(default_tab_stop);
                // Tab positions are authored: never move text across a tab stop.
                compression = 0.0;
                has_tab = true;
            }
            let capacity = if has_tab || !compress_punctuation {
                0.0
            } else {
                punctuation_compression(character, advance)
            };
            if end > start
                && !is_line_start_prohibited(character)
                && !fits_compressed(
                    used,
                    advance,
                    width,
                    compression + capacity,
                    is_closing_punctuation(character),
                )
            {
                // A Latin word may use its preceding space. CJK has a break
                // opportunity between characters, so an earlier space is irrelevant.
                if character.is_ascii_alphanumeric()
                    && characters[end - 1].0.is_ascii_alphanumeric()
                    && let Some(space) = whitespace
                {
                    end = space + 1;
                }
                break;
            }
            used += advance;
            compression += capacity;
            if matches!(character, ' ' | '\t') {
                whitespace = Some(end);
            }
            end += 1;
        }
        breaks.push(end);
        start = end;
        width = following_width.max(1.0);
    }
    if breaks.is_empty() {
        breaks.push(0);
    }
    breaks
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn closing_punctuation_never_starts_a_soft_wrapped_line() {
        for closing in "、。，．！？：；）］】》〉」』〕”’,.;:!?)]}…".chars()
        {
            for compress in [false, true] {
                let characters = [('中', 16.0), ('文', 16.0), (closing, 16.0), ('后', 16.0)];
                assert_eq!(
                    horizontal_line_breaks(&characters, 32.0, 32.0, 48.0, compress),
                    vec![3, 4]
                );
            }
        }
        assert_eq!(
            horizontal_line_breaks(
                &[('中', 16.0), ('\n', 0.0), ('。', 16.0)],
                16.0,
                16.0,
                48.0,
                false
            ),
            vec![2, 3]
        );
    }

    #[test]
    fn fits_supplied_chinese_line_without_shrinking_letters_or_losing_text() {
        let text = "中去除个人信息的行为，使其保持不可被检索、访问的状态。";
        let characters: Vec<_> = text.chars().map(|c| (c, 64.0 / 3.0)).collect();
        assert_eq!(
            compressed_line_breaks(&characters, 553.7333, 553.7333, 28.0),
            vec![27]
        );
        assert_eq!(
            horizontal_line_breaks(&characters, 553.7333, 553.7333, 28.0, false),
            vec![25, 27],
            "formats without punctuation compression must retain natural advances"
        );
        assert_eq!(punctuation_compression('中', 20.0), 0.0);
        assert_eq!(punctuation_compression(',', 10.0), 0.0);
        assert_eq!(punctuation_compression('，', 20.0), 10.0);
        assert_eq!(
            compressed_line_breaks(&[('中', 20.0); 3], 40.0, 40.0, 28.0),
            vec![2, 3]
        );
    }
}
