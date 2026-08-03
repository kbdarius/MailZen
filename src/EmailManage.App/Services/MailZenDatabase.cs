using EmailManage.Models;
using Microsoft.Data.Sqlite;
using System.IO;

namespace EmailManage.Services;

/// <summary>
/// Owns the local MailZen index schema and transactional message writes.
/// Email content is intentionally never sent to DiagnosticLogger.
/// </summary>
public sealed class MailZenDatabase
{
    public const int CurrentSchemaVersion = 2;
    public string DatabasePath { get; }
    private string ConnectionString => new SqliteConnectionStringBuilder
    {
        DataSource = DatabasePath,
        Mode = SqliteOpenMode.ReadWriteCreate,
        Cache = SqliteCacheMode.Shared
    }.ToString();

    public MailZenDatabase(string? databasePath = null)
    {
        DatabasePath = databasePath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MailZen", "MailZen.db");
    }

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqliteTransaction)await connection.BeginTransactionAsync(cancellationToken);

        await ExecuteAsync(connection, transaction, "PRAGMA foreign_keys = ON;", cancellationToken);
        await ExecuteAsync(connection, transaction, """
            CREATE TABLE IF NOT EXISTS schema_version (
                version INTEGER NOT NULL,
                applied_utc TEXT NOT NULL
            );
            """, cancellationToken);

        var version = await ScalarIntAsync(connection, transaction,
            "SELECT COALESCE(MAX(version), 0) FROM schema_version;", cancellationToken);
        if (version < 1)
        {
            await ExecuteAsync(connection, transaction, """
                CREATE TABLE accounts (
                    id TEXT PRIMARY KEY,
                    store_id TEXT NOT NULL UNIQUE,
                    display_name TEXT NOT NULL,
                    email_address TEXT NOT NULL,
                    provider_hint TEXT,
                    is_enabled INTEGER NOT NULL DEFAULT 1,
                    last_seen_utc TEXT
                );
                CREATE TABLE folders (
                    id TEXT PRIMARY KEY,
                    account_id TEXT NOT NULL REFERENCES accounts(id),
                    entry_id TEXT NOT NULL,
                    folder_path TEXT NOT NULL,
                    folder_type TEXT NOT NULL,
                    is_enabled INTEGER NOT NULL DEFAULT 1,
                    UNIQUE(account_id, entry_id)
                );
                CREATE TABLE messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    message_id TEXT NOT NULL UNIQUE,
                    account_id TEXT NOT NULL REFERENCES accounts(id),
                    folder_id TEXT NOT NULL REFERENCES folders(id),
                    store_id TEXT NOT NULL,
                    entry_id TEXT NOT NULL,
                    internet_message_id TEXT,
                    conversation_id TEXT,
                    subject TEXT NOT NULL DEFAULT '',
                    sender_name TEXT NOT NULL DEFAULT '',
                    sender_address TEXT NOT NULL DEFAULT '',
                    to_recipients TEXT NOT NULL DEFAULT '',
                    cc_recipients TEXT NOT NULL DEFAULT '',
                    received_utc TEXT NOT NULL,
                    sent_utc TEXT,
                    is_unread INTEGER NOT NULL DEFAULT 0,
                    importance INTEGER NOT NULL DEFAULT 0,
                    has_attachments INTEGER NOT NULL DEFAULT 0,
                    attachment_names TEXT NOT NULL DEFAULT '',
                    body_text TEXT NOT NULL DEFAULT '',
                    body_hash TEXT,
                    source_modified_utc TEXT,
                    indexed_utc TEXT NOT NULL,
                    last_seen_utc TEXT NOT NULL,
                    is_missing INTEGER NOT NULL DEFAULT 0
                );
                CREATE TABLE identifier_aliases (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    message_id INTEGER NOT NULL REFERENCES messages(id),
                    store_id TEXT NOT NULL,
                    entry_id TEXT NOT NULL,
                    created_utc TEXT NOT NULL,
                    UNIQUE(store_id, entry_id)
                );
                CREATE TABLE sync_state (
                    account_id TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    last_successful_run_utc TEXT,
                    last_attempted_run_utc TEXT,
                    high_water_received_utc TEXT,
                    error_summary TEXT,
                    consecutive_failure_count INTEGER NOT NULL DEFAULT 0,
                    PRIMARY KEY(account_id, folder_id)
                );
                CREATE TABLE search_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    search_text TEXT NOT NULL,
                    selected_scope TEXT NOT NULL,
                    model_profile TEXT,
                    duration_ms INTEGER,
                    result_count INTEGER,
                    created_utc TEXT NOT NULL
                );
                CREATE INDEX ix_messages_account_received ON messages(account_id, received_utc DESC);
                CREATE INDEX ix_messages_folder_received ON messages(folder_id, received_utc DESC);
                CREATE INDEX ix_messages_internet_id ON messages(internet_message_id);
                CREATE VIRTUAL TABLE messages_fts USING fts5(
                    subject, sender_name, sender_address, to_recipients, cc_recipients,
                    body_text, attachment_names, content='messages', content_rowid='id'
                );
                CREATE TRIGGER messages_ai AFTER INSERT ON messages BEGIN
                    INSERT INTO messages_fts(rowid, subject, sender_name, sender_address, to_recipients, cc_recipients, body_text, attachment_names)
                    VALUES (new.id, new.subject, new.sender_name, new.sender_address, new.to_recipients, new.cc_recipients, new.body_text, new.attachment_names);
                END;
                CREATE TRIGGER messages_au AFTER UPDATE ON messages BEGIN
                    INSERT INTO messages_fts(messages_fts, rowid, subject, sender_name, sender_address, to_recipients, cc_recipients, body_text, attachment_names)
                    VALUES ('delete', old.id, old.subject, old.sender_name, old.sender_address, old.to_recipients, old.cc_recipients, old.body_text, old.attachment_names);
                    INSERT INTO messages_fts(rowid, subject, sender_name, sender_address, to_recipients, cc_recipients, body_text, attachment_names)
                    VALUES (new.id, new.subject, new.sender_name, new.sender_address, new.to_recipients, new.cc_recipients, new.body_text, new.attachment_names);
                END;
                CREATE TRIGGER messages_ad AFTER DELETE ON messages BEGIN
                    INSERT INTO messages_fts(messages_fts, rowid, subject, sender_name, sender_address, to_recipients, cc_recipients, body_text, attachment_names)
                    VALUES ('delete', old.id, old.subject, old.sender_name, old.sender_address, old.to_recipients, old.cc_recipients, old.body_text, old.attachment_names);
                END;
                INSERT INTO schema_version(version, applied_utc) VALUES (1, $now);
                """, cancellationToken, ("$now", (object)DateTime.UtcNow.ToString("O")));
        }

        if (version < 2)
        {
            await ExecuteAsync(connection, transaction, """
                CREATE TABLE IF NOT EXISTS sync_coverage (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id TEXT NOT NULL REFERENCES accounts(id),
                    folder_id TEXT NOT NULL REFERENCES folders(id),
                    range_start_utc TEXT NOT NULL,
                    range_end_utc TEXT NOT NULL,
                    completed_utc TEXT NOT NULL,
                    status TEXT NOT NULL,
                    error_summary TEXT,
                    UNIQUE(account_id, folder_id, range_start_utc, range_end_utc)
                );
                CREATE INDEX IF NOT EXISTS ix_sync_coverage_account_range
                    ON sync_coverage(account_id, range_start_utc, range_end_utc);
                INSERT INTO schema_version(version, applied_utc) VALUES (2, $now);
                """, cancellationToken, ("$now", (object)DateTime.UtcNow.ToString("O")));
        }

        await transaction.CommitAsync(cancellationToken);
    }

    public async Task UpsertMessageAsync(IndexedMessage message, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO accounts(id, store_id, display_name, email_address)
            VALUES ($account, $store, $account, '')
            ON CONFLICT(id) DO UPDATE SET store_id = excluded.store_id;
            INSERT INTO folders(id, account_id, entry_id, folder_path, folder_type)
            VALUES ($folder, $account, $folder, $folder, 'unknown')
            ON CONFLICT(id) DO NOTHING;
            INSERT INTO messages(message_id, account_id, folder_id, store_id, entry_id, internet_message_id,
                conversation_id, subject, sender_name, sender_address, received_utc, is_unread,
                has_attachments, attachment_names, body_text, indexed_utc, last_seen_utc, is_missing)
            VALUES ($message, $account, $folder, $store, $entry, $internet, $conversation, $subject,
                $senderName, $senderAddress, $received, $unread, $attachments, $attachmentNames, $body,
                $now, $now, 0)
            ON CONFLICT(message_id) DO UPDATE SET
                account_id = excluded.account_id, folder_id = excluded.folder_id, store_id = excluded.store_id,
                entry_id = excluded.entry_id, internet_message_id = excluded.internet_message_id,
                conversation_id = excluded.conversation_id, subject = excluded.subject,
                sender_name = excluded.sender_name, sender_address = excluded.sender_address,
                received_utc = excluded.received_utc, is_unread = excluded.is_unread,
                has_attachments = excluded.has_attachments, attachment_names = excluded.attachment_names,
                body_text = excluded.body_text, last_seen_utc = excluded.last_seen_utc, is_missing = 0;
            """;
        command.Parameters.AddWithValue("$account", message.AccountId);
        command.Parameters.AddWithValue("$store", message.StoreId);
        command.Parameters.AddWithValue("$folder", message.FolderId);
        command.Parameters.AddWithValue("$message", message.MessageId);
        command.Parameters.AddWithValue("$entry", message.EntryId);
        command.Parameters.AddWithValue("$internet", (object?)message.InternetMessageId ?? DBNull.Value);
        command.Parameters.AddWithValue("$conversation", (object?)message.ConversationId ?? DBNull.Value);
        command.Parameters.AddWithValue("$subject", message.Subject);
        command.Parameters.AddWithValue("$senderName", message.SenderName);
        command.Parameters.AddWithValue("$senderAddress", message.SenderAddress);
        command.Parameters.AddWithValue("$received", message.ReceivedUtc.ToUniversalTime().ToString("O"));
        command.Parameters.AddWithValue("$unread", message.IsUnread ? 1 : 0);
        command.Parameters.AddWithValue("$attachments", message.HasAttachments ? 1 : 0);
        command.Parameters.AddWithValue("$attachmentNames", message.AttachmentNames);
        command.Parameters.AddWithValue("$body", message.BodyText);
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task UpsertMessagesAsync(IReadOnlyList<IndexedMessage> messages, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqliteTransaction)await connection.BeginTransactionAsync(cancellationToken);
        await using (var accountFolderCommand = connection.CreateCommand())
        {
            accountFolderCommand.Transaction = transaction;
            accountFolderCommand.CommandText = """
                INSERT INTO accounts(id, store_id, display_name, email_address) VALUES ($account, $store, $account, '')
                ON CONFLICT(id) DO UPDATE SET store_id = excluded.store_id;
                INSERT INTO folders(id, account_id, entry_id, folder_path, folder_type) VALUES ($folder, $account, $folder, $folder, 'unknown')
                ON CONFLICT(id) DO NOTHING;
                """;
            var account = accountFolderCommand.Parameters.Add("$account", SqliteType.Text);
            var store = accountFolderCommand.Parameters.Add("$store", SqliteType.Text);
            var folder = accountFolderCommand.Parameters.Add("$folder", SqliteType.Text);
            accountFolderCommand.Prepare();
            foreach (var message in messages.GroupBy(m => (m.AccountId, m.StoreId, m.FolderId)).Select(g => g.First()))
            {
                account.Value = message.AccountId; store.Value = message.StoreId; folder.Value = message.FolderId;
                await accountFolderCommand.ExecuteNonQueryAsync(cancellationToken);
            }
        }
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO messages(message_id, account_id, folder_id, store_id, entry_id, internet_message_id, conversation_id,
                subject, sender_name, sender_address, received_utc, is_unread, has_attachments, attachment_names, body_text,
                indexed_utc, last_seen_utc, is_missing)
            VALUES ($message, $account, $folder, $store, $entry, $internet, $conversation, $subject, $senderName, $senderAddress,
                $received, $unread, $attachments, $attachmentNames, $body, $now, $now, 0)
            ON CONFLICT(message_id) DO UPDATE SET account_id = excluded.account_id, folder_id = excluded.folder_id,
                store_id = excluded.store_id, entry_id = excluded.entry_id, internet_message_id = excluded.internet_message_id,
                subject = excluded.subject, sender_name = excluded.sender_name, sender_address = excluded.sender_address,
                received_utc = excluded.received_utc, is_unread = excluded.is_unread, has_attachments = excluded.has_attachments,
                attachment_names = excluded.attachment_names, body_text = excluded.body_text, last_seen_utc = excluded.last_seen_utc, is_missing = 0;
            """;
        var accountParameter = command.Parameters.Add("$account", SqliteType.Text);
        var storeParameter = command.Parameters.Add("$store", SqliteType.Text);
        var folderParameter = command.Parameters.Add("$folder", SqliteType.Text);
        var messageParameter = command.Parameters.Add("$message", SqliteType.Text);
        var entryParameter = command.Parameters.Add("$entry", SqliteType.Text);
        var internetParameter = command.Parameters.Add("$internet", SqliteType.Text);
        var conversationParameter = command.Parameters.Add("$conversation", SqliteType.Text);
        var subjectParameter = command.Parameters.Add("$subject", SqliteType.Text);
        var senderNameParameter = command.Parameters.Add("$senderName", SqliteType.Text);
        var senderAddressParameter = command.Parameters.Add("$senderAddress", SqliteType.Text);
        var receivedParameter = command.Parameters.Add("$received", SqliteType.Text);
        var unreadParameter = command.Parameters.Add("$unread", SqliteType.Integer);
        var attachmentsParameter = command.Parameters.Add("$attachments", SqliteType.Integer);
        var attachmentNamesParameter = command.Parameters.Add("$attachmentNames", SqliteType.Text);
        var bodyParameter = command.Parameters.Add("$body", SqliteType.Text);
        var nowParameter = command.Parameters.Add("$now", SqliteType.Text);
        command.Prepare();
        foreach (var message in messages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            accountParameter.Value = message.AccountId; storeParameter.Value = message.StoreId; folderParameter.Value = message.FolderId;
            messageParameter.Value = message.MessageId; entryParameter.Value = message.EntryId;
            internetParameter.Value = (object?)message.InternetMessageId ?? DBNull.Value; conversationParameter.Value = (object?)message.ConversationId ?? DBNull.Value;
            subjectParameter.Value = message.Subject; senderNameParameter.Value = message.SenderName; senderAddressParameter.Value = message.SenderAddress;
            receivedParameter.Value = message.ReceivedUtc.ToUniversalTime().ToString("O"); unreadParameter.Value = message.IsUnread ? 1 : 0;
            attachmentsParameter.Value = message.HasAttachments ? 1 : 0; attachmentNamesParameter.Value = message.AttachmentNames;
            bodyParameter.Value = message.BodyText; nowParameter.Value = DateTime.UtcNow.ToString("O");
            await command.ExecuteNonQueryAsync(cancellationToken);
        }
        await transaction.CommitAsync(cancellationToken);
    }

    public async Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        if (request.Scope.AccountIds.Count == 0 || string.IsNullOrWhiteSpace(request.Query)) return Array.Empty<SearchResult>();
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();

        var accountParameters = request.Scope.AccountIds.Select((id, i) =>
        {
            var name = "$account" + i;
            command.Parameters.AddWithValue(name, id);
            return name;
        }).ToArray();
        var match = request.Mode == SearchMode.Boolean
            ? SearchQueryText.BuildBooleanQuery(request.Query)
            : string.Join(" AND ", SearchQueryText.BuildLocalQuery(request.Query)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(term => $"\"{term.Replace("\"", "\"\"")}\"*"));
        if (string.IsNullOrWhiteSpace(match)) return Array.Empty<SearchResult>();
        command.Parameters.AddWithValue("$match", match);
        command.Parameters.AddWithValue("$limit", Math.Clamp(request.MaxResults, 1, 250));

        var predicates = new List<string> { $"m.account_id IN ({string.Join(",", accountParameters)})" };
        if (request.Scope.FolderIds is { Count: > 0 })
        {
            var folderParameters = request.Scope.FolderIds.Select((id, i) => { var name = "$folder" + i; command.Parameters.AddWithValue(name, id); return name; }).ToArray();
            predicates.Add($"m.folder_id IN ({string.Join(",", folderParameters)})");
        }
        if (request.Scope.ReceivedAfterUtc.HasValue) { command.Parameters.AddWithValue("$after", request.Scope.ReceivedAfterUtc.Value.ToUniversalTime().ToString("O")); predicates.Add("m.received_utc >= $after"); }
        if (request.Scope.ReceivedBeforeUtc.HasValue) { command.Parameters.AddWithValue("$before", request.Scope.ReceivedBeforeUtc.Value.ToUniversalTime().ToString("O")); predicates.Add("m.received_utc <= $before"); }
        if (request.Scope.IsUnread.HasValue) { command.Parameters.AddWithValue("$unread", request.Scope.IsUnread.Value ? 1 : 0); predicates.Add("m.is_unread = $unread"); }
        if (request.Scope.HasAttachments.HasValue) { command.Parameters.AddWithValue("$attachments", request.Scope.HasAttachments.Value ? 1 : 0); predicates.Add("m.has_attachments = $attachments"); }

        command.CommandText = $"""
            SELECT m.message_id, m.account_id, m.folder_id, m.store_id, m.entry_id, m.internet_message_id,
                   m.conversation_id, m.subject, m.sender_name, m.sender_address, m.body_text,
                   m.received_utc, m.is_unread, m.has_attachments, m.attachment_names,
                   bm25(messages_fts) AS score,
                   snippet(messages_fts, 5, '<b>', '</b>', '…', 24) AS excerpt
            FROM messages_fts
            JOIN messages m ON m.id = messages_fts.rowid
            WHERE messages_fts MATCH $match AND {string.Join(" AND ", predicates)}
            ORDER BY score ASC, m.received_utc DESC
            LIMIT $limit;
            """;

        var results = new List<SearchResult>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var message = new IndexedMessage(
                reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), reader.GetString(4),
                reader.IsDBNull(5) ? null : reader.GetString(5), reader.GetString(7), reader.GetString(8), reader.GetString(9),
                reader.GetString(10), DateTime.Parse(reader.GetString(11)).ToUniversalTime(), reader.GetInt32(12) != 0,
                reader.GetInt32(13) != 0, reader.GetString(14), reader.IsDBNull(6) ? null : reader.GetString(6));
            results.Add(new SearchResult(message, reader.GetDouble(15), reader.IsDBNull(16) ? message.BodyText[..Math.Min(240, message.BodyText.Length)] : reader.GetString(16)));
        }
        return results;
    }

    public async Task<IndexCoverage> GetIndexCoverageAsync(IReadOnlySet<string> accountIds, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        if (accountIds.Count == 0) return new IndexCoverage(0, null, null);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        var parameters = accountIds.Select((id, index) =>
        {
            var name = "$account" + index;
            command.Parameters.AddWithValue(name, id);
            return name;
        }).ToArray();
        command.CommandText = $"SELECT COUNT(*), MIN(received_utc), MAX(received_utc) FROM messages WHERE account_id IN ({string.Join(",", parameters)}) AND is_missing = 0;";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        await reader.ReadAsync(cancellationToken);
        return new IndexCoverage(
            reader.GetInt32(0),
            reader.IsDBNull(1) ? null : DateTime.Parse(reader.GetString(1)).ToUniversalTime(),
            reader.IsDBNull(2) ? null : DateTime.Parse(reader.GetString(2)).ToUniversalTime());
    }

    public async Task<IReadOnlyDictionary<string, IndexCoverage>> GetIndexCoverageByAccountAsync(
        IReadOnlySet<string> accountIds, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        var result = accountIds.ToDictionary(id => id, _ => new IndexCoverage(0, null, null));
        if (accountIds.Count == 0) return result;
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        var parameters = accountIds.Select((id, index) =>
        {
            var name = "$account" + index;
            command.Parameters.AddWithValue(name, id);
            return name;
        }).ToArray();
        command.CommandText = $"SELECT account_id, COUNT(*), MIN(received_utc), MAX(received_utc) FROM messages WHERE account_id IN ({string.Join(",", parameters)}) AND is_missing = 0 GROUP BY account_id;";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            result[reader.GetString(0)] = new IndexCoverage(
                reader.GetInt32(1),
                reader.IsDBNull(2) ? null : DateTime.Parse(reader.GetString(2)).ToUniversalTime(),
                reader.IsDBNull(3) ? null : DateTime.Parse(reader.GetString(3)).ToUniversalTime());
        }
        return result;
    }

    public async Task RecordSyncCoverageAsync(string accountId, string folderId, DateTime rangeStartUtc,
        DateTime rangeEndUtc, DateTime completedUtc, string status = "complete", string? error = null,
        CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO sync_coverage(account_id, folder_id, range_start_utc, range_end_utc, completed_utc, status, error_summary)
            VALUES ($account, $folder, $start, $end, $completed, $status, $error)
            ON CONFLICT(account_id, folder_id, range_start_utc, range_end_utc) DO UPDATE SET
                completed_utc = excluded.completed_utc,
                status = excluded.status,
                error_summary = excluded.error_summary;
            """;
        command.Parameters.AddWithValue("$account", accountId);
        command.Parameters.AddWithValue("$folder", folderId);
        command.Parameters.AddWithValue("$start", rangeStartUtc.ToUniversalTime().ToString("O"));
        command.Parameters.AddWithValue("$end", rangeEndUtc.ToUniversalTime().ToString("O"));
        command.Parameters.AddWithValue("$completed", completedUtc.ToUniversalTime().ToString("O"));
        command.Parameters.AddWithValue("$status", status);
        command.Parameters.AddWithValue("$error", (object?)error ?? DBNull.Value);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<DateTime> GetEffectiveSyncStartUtcAsync(string accountId, string folderId,
        DateTime requestedStartUtc, DateTime requestedEndUtc, TimeSpan safetyOverlap,
        CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        var requestedStart = requestedStartUtc.ToUniversalTime();
        var requestedEnd = requestedEndUtc.ToUniversalTime();
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT range_start_utc, range_end_utc
            FROM sync_coverage
            WHERE account_id = $account AND folder_id = $folder AND status = 'complete'
              AND range_end_utc > $start AND range_start_utc < $end
            ORDER BY range_start_utc;
            """;
        command.Parameters.AddWithValue("$account", accountId);
        command.Parameters.AddWithValue("$folder", folderId);
        command.Parameters.AddWithValue("$start", requestedStart.ToString("O"));
        command.Parameters.AddWithValue("$end", requestedEnd.ToString("O"));

        var cursor = requestedStart;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var rangeStart = DateTime.Parse(reader.GetString(0)).ToUniversalTime();
            var rangeEnd = DateTime.Parse(reader.GetString(1)).ToUniversalTime();
            if (rangeEnd <= cursor) continue;
            if (rangeStart > cursor)
                return cursor > requestedStart ? cursor.Subtract(safetyOverlap) : requestedStart;
            cursor = rangeEnd > cursor ? rangeEnd : cursor;
            if (cursor >= requestedEnd)
                return requestedEnd.Subtract(safetyOverlap) > requestedStart
                    ? requestedEnd.Subtract(safetyOverlap)
                    : requestedStart;
        }

        return cursor > requestedStart ? cursor.Subtract(safetyOverlap) : requestedStart;
    }

    public async Task<DateTime?> GetLatestSuccessfulSyncUtcAsync(IReadOnlySet<string> accountIds, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        if (accountIds.Count == 0) return null;
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        var parameters = accountIds.Select((id, index) =>
        {
            var name = "$account" + index;
            command.Parameters.AddWithValue(name, id);
            return name;
        }).ToArray();
        command.CommandText = $"SELECT MAX(last_successful_run_utc) FROM sync_state WHERE account_id IN ({string.Join(",", parameters)});";
        var value = await command.ExecuteScalarAsync(cancellationToken);
        return value is null or DBNull ? null : DateTime.Parse((string)value).ToUniversalTime();
    }

    public async Task<DateTime?> GetEarliestSuccessfulSyncUtcAsync(IReadOnlySet<string> accountIds, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        if (accountIds.Count == 0) return null;
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        var parameters = accountIds.Select((id, index) =>
        {
            var name = "$account" + index;
            command.Parameters.AddWithValue(name, id);
            return name;
        }).ToArray();
        command.CommandText = $"SELECT MIN(last_successful_run_utc) FROM sync_state WHERE account_id IN ({string.Join(",", parameters)}) AND last_successful_run_utc IS NOT NULL;";
        var value = await command.ExecuteScalarAsync(cancellationToken);
        return value is null or DBNull ? null : DateTime.Parse((string)value).ToUniversalTime();
    }

    public async Task MarkFolderSyncSuccessAsync(string accountId, string folderId, DateTime completedUtc, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO sync_state(account_id, folder_id, last_successful_run_utc, last_attempted_run_utc, consecutive_failure_count)
            VALUES ($account, $folder, $completed, $completed, 0)
            ON CONFLICT(account_id, folder_id) DO UPDATE SET
                last_successful_run_utc = excluded.last_successful_run_utc,
                last_attempted_run_utc = excluded.last_attempted_run_utc,
                error_summary = NULL,
                consecutive_failure_count = 0;
            """;
        command.Parameters.AddWithValue("$account", accountId);
        command.Parameters.AddWithValue("$folder", folderId);
        command.Parameters.AddWithValue("$completed", completedUtc.ToUniversalTime().ToString("O"));
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task MarkFolderSyncFailureAsync(string accountId, string folderId, DateTime attemptedUtc, string error, CancellationToken cancellationToken = default)
    {
        await InitializeAsync(cancellationToken);
        await using var connection = new SqliteConnection(ConnectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO sync_state(account_id, folder_id, last_attempted_run_utc, error_summary, consecutive_failure_count)
            VALUES ($account, $folder, $attempted, $error, 1)
            ON CONFLICT(account_id, folder_id) DO UPDATE SET
                last_attempted_run_utc = excluded.last_attempted_run_utc,
                error_summary = excluded.error_summary,
                consecutive_failure_count = sync_state.consecutive_failure_count + 1;
            """;
        command.Parameters.AddWithValue("$account", accountId);
        command.Parameters.AddWithValue("$folder", folderId);
        command.Parameters.AddWithValue("$attempted", attemptedUtc.ToUniversalTime().ToString("O"));
        command.Parameters.AddWithValue("$error", error.Length > 500 ? error[..500] : error);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task ExecuteAsync(SqliteConnection connection, SqliteTransaction transaction, string sql, CancellationToken token, params (string Name, object Value)[] parameters)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        foreach (var (name, value) in parameters) command.Parameters.AddWithValue(name, value);
        await command.ExecuteNonQueryAsync(token);
    }

    private static async Task<int> ScalarIntAsync(SqliteConnection connection, SqliteTransaction transaction, string sql, CancellationToken token)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        return Convert.ToInt32(await command.ExecuteScalarAsync(token));
    }
}
