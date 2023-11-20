namespace ispring_converter;

public record Question(
    ushort Number,
    string Text,
    bool IsMultiply,
    byte Points,
    byte Attempt,
    string MessageIfCorrect,
    string MessageIfIncorrect,
    Answer[] Answers);

public record Answer(
    bool IsValid,
    string Text
);