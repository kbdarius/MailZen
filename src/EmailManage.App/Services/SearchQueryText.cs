using System.Text.RegularExpressions;
using EmailManage.Models;

namespace EmailManage.Services;

public static partial class SearchQueryText
{
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "an", "and", "are", "can", "could", "email", "emails", "find", "for", "from", "get",
        "got", "help", "i", "im", "i'm", "in", "is", "it", "looking", "me", "message", "messages",
        "of", "please", "search", "that", "the", "this", "to", "was", "were", "with", "you"
    };

    public static string BuildLocalQuery(string query)
    {
        var terms = Tokenize(query)
            .Where(term => !StopWords.Contains(term))
            .Select(NormalizeTerm)
            .Where(term => term.Length >= 2)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return string.Join(' ', terms);
    }

    public static string BuildIntentQuery(SearchIntent intent)
    {
        var terms = intent.People
            .Concat(intent.Organizations)
            .Concat(intent.RequiredKeywords)
            .Concat(intent.OptionalKeywords)
            .SelectMany(Tokenize)
            .Select(NormalizeTerm)
            .Where(term => term.Length >= 2)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return string.Join(' ', terms);
    }

    public static string BuildBooleanQuery(string query)
    {
        var output = new List<string>();
        foreach (var token in BooleanTokenRegex().Matches(query).Select(match => match.Value))
        {
            if (token is "(" or ")") { output.Add(token); continue; }
            if (token is "&" or "AND") { output.Add("AND"); continue; }
            if (token is "|" or "OR") { output.Add("OR"); continue; }
            if (token.Equals("NOT", StringComparison.OrdinalIgnoreCase)) { output.Add("NOT"); continue; }
            var term = NormalizeTerm(token.ToLowerInvariant());
            if (term.Length >= 2) output.Add($"\"{term.Replace("\"", "\"\"")}\"*");
        }
        return string.Join(' ', output).Replace("( ", "(").Replace(" )", ")");
    }

    private static IEnumerable<string> Tokenize(string value) =>
        TokenRegex().Matches(value).Select(match => match.Value.ToLowerInvariant());

    private static string NormalizeTerm(string term) => term switch
    {
        "qoute" or "qout" => "quote",
        "moveing" => "moving",
        _ => term
    };

    [GeneratedRegex("[\\p{L}\\p{N}']+")]
    private static partial Regex TokenRegex();

    [GeneratedRegex("\\(|\\)|&|\\||\\bAND\\b|\\bOR\\b|\\bNOT\\b|[\\p{L}\\p{N}']+", RegexOptions.IgnoreCase)]
    private static partial Regex BooleanTokenRegex();
}
