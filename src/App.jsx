import { useState } from "react";
import "./App.css";

const API_URL = "https://guestinvitemobsa-a3g2dxbdaqfuazap.australiasoutheast-01.azurewebsites.net/api/inviteguest";
const API_ORIGIN = new URL(API_URL).origin;
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_NAME_LENGTH = 40;
const MAX_BULK_INVITES = 100;
const DEFAULT_REDIRECT_URL = "https://myapplications.microsoft.com";

function isValidHttpUrl(value) {
  if (!value) return false;

  try {
    const url = new URL(value);
    const isHttp = url.protocol === "https:" || url.protocol === "http:";
    const isLocalhost = ["localhost", "127.0.0.1"].includes(url.hostname);
    if (!isHttp) return false;
    if (url.protocol === "http:" && !isLocalhost) return false;
    return true;
  } catch {
    return false;
  }
}

function sanitizeName(value) {
  return value.trim().replace(/\s+/g, " ").slice(0, MAX_NAME_LENGTH);
}

function normalizeEmail(value) {
  return value.trim().toLowerCase();
}

function getDisplayName(firstName, lastName) {
  return [firstName, lastName].filter(Boolean).join(" ").trim();
}

function parseBoolean(value, fallback = true) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "true") return true;
  if (normalized === "false") return false;
  return fallback;
}

function safeResponse(payload) {
  if (!payload || typeof payload !== "object") return null;

  const safe = {};
  if (typeof payload.message === "string") safe.message = payload.message;
  if (typeof payload.status === "string") safe.status = payload.status;
  if (typeof payload.code === "string") safe.code = payload.code;
  if (typeof payload.error === "string") safe.error = payload.error;
  if (typeof payload.detail === "string") safe.detail = payload.detail;
  return Object.keys(safe).length ? safe : null;
}

function buildInvitePayload(payload) {
  const body = {
    email: payload.email,
    displayName: payload.displayName || "",
  };

  if (payload.firstName) body.firstName = payload.firstName;
  if (payload.lastName) body.lastName = payload.lastName;
  if (payload.inviteRedirectUrl) {
    body.inviteRedirectUrl = payload.inviteRedirectUrl;
    body.inviteRedirectURL = payload.inviteRedirectUrl;
  }
  if (typeof payload.sendEmail === "boolean") body.sendEmail = payload.sendEmail;
  if (payload.customizedMessageBody) body.customizedMessageBody = payload.customizedMessageBody;
  if (typeof payload.resetRedemption === "boolean") body.resetRedemption = payload.resetRedemption;

  return body;
}

function getRedeemUrl(result) {
  const url = result?.data?.inviteRedeemUrl;
  return typeof url === "string" && url.startsWith("http") ? url : null;
}

function parseCsvLine(line) {
  const out = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];

    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      out.push(current.trim());
      current = "";
      continue;
    }

    current += ch;
  }

  out.push(current.trim());
  return out;
}

function parseManualBulk(rawText, defaults) {
  const lines = rawText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const entries = [];
  const errors = [];

  lines.forEach((line, index) => {
    const [emailRaw, firstRaw = "", lastRaw = ""] = parseCsvLine(line);
    const email = normalizeEmail(emailRaw);
    const firstName = sanitizeName(firstRaw);
    const lastName = sanitizeName(lastRaw);

    if (!EMAIL_REGEX.test(email)) {
      errors.push(`Line ${index + 1}: invalid email`);
      return;
    }

    entries.push({
      email,
      firstName,
      lastName,
      displayName: getDisplayName(firstName, lastName),
      inviteRedirectUrl: defaults.redirectUrl,
      sendEmail: defaults.sendEmail,
      customizedMessageBody: defaults.customizedMessageBody,
      resetRedemption: defaults.resetRedemption,
    });
  });

  return { entries, errors };
}

function parseEntraTemplateCsv(rawText, defaults = { resetRedemption: false }) {
  const normalized = rawText.replace(/^\uFEFF/, "");
  const lines = normalized
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2) {
    return {
      entries: [],
      errors: ["CSV is invalid. Expected version row and header row."],
    };
  }

  const headers = parseCsvLine(lines[1]);
  const emailIdx = headers.findIndex((h) => h.includes("[inviteeEmail]"));
  const redirectIdx = headers.findIndex((h) => h.includes("[inviteRedirectURL]"));
  const sendEmailIdx = headers.findIndex((h) => h.includes("[sendEmail]"));
  const messageIdx = headers.findIndex((h) => h.includes("[customizedMessageBody]"));

  if (emailIdx < 0 || redirectIdx < 0) {
    return {
      entries: [],
      errors: ["CSV missing required columns: [inviteeEmail] and/or [inviteRedirectURL]."],
    };
  }

  const entries = [];
  const errors = [];

  for (let i = 2; i < lines.length; i += 1) {
    const row = parseCsvLine(lines[i]);
    const rawEmail = (row[emailIdx] || "").replace(/^Example:\s*/i, "");
    if (!rawEmail || rawEmail.startsWith("Example:")) continue;

    const email = normalizeEmail(rawEmail);
    const redirect = (row[redirectIdx] || "").trim();

    if (!EMAIL_REGEX.test(email)) {
      errors.push(`Line ${i + 1}: invalid inviteeEmail`);
      continue;
    }

    if (!isValidHttpUrl(redirect)) {
      errors.push(`Line ${i + 1}: invalid inviteRedirectURL`);
      continue;
    }

    entries.push({
      email,
      firstName: "",
      lastName: "",
      displayName: "",
      inviteRedirectUrl: redirect,
      sendEmail: parseBoolean(sendEmailIdx >= 0 ? row[sendEmailIdx] : true, true),
      customizedMessageBody: messageIdx >= 0 ? (row[messageIdx] || "").trim() : "",
      resetRedemption: defaults.resetRedemption,
    });
  }

  return { entries, errors };
}

async function sendInvite(apiUrl, payload) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 12000);

  try {
    const res = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildInvitePayload(payload)),
      credentials: "include",
      signal: controller.signal,
    });

    const text = await res.text();
    let parsed = null;
    try {
      parsed = text ? JSON.parse(text) : null;
    } catch {
      parsed = null;
    }

    const payloadSafe = safeResponse(parsed);

    if (!res.ok) {
      const fallbackText = text ? text.slice(0, 280) : null;
      return {
        ok: false,
        data: {
          ...(payloadSafe || {}),
          error:
            (payloadSafe && payloadSafe.error) ||
            `Invite failed with HTTP ${res.status}. Check API permissions and payload format.`,
          ...(fallbackText ? { backendResponse: fallbackText } : {}),
        },
      };
    }

    return { ok: true, data: payloadSafe || { message: "Invite sent." } };
  } catch (err) {
    return {
      ok: false,
      data: {
        error:
          err?.name === "AbortError"
            ? "Request timed out. Try again."
            : "Network request failed. Check connection and API URL.",
      },
    };
  } finally {
    clearTimeout(timeout);
  }
}

export default function App() {
  const [mode, setMode] = useState("single");

  const [email, setEmail] = useState("");
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");

  const [redirectUrl, setRedirectUrl] = useState(DEFAULT_REDIRECT_URL);
  const [sendEmailFlag, setSendEmailFlag] = useState(true);
  const [resetRedemptionFlag, setResetRedemptionFlag] = useState(false);
  const [customizedMessageBody, setCustomizedMessageBody] = useState("");

  const [bulkInput, setBulkInput] = useState("");
  const [bulkEntriesFromFile, setBulkEntriesFromFile] = useState([]);
  const [bulkFileName, setBulkFileName] = useState("");

  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [authState, setAuthState] = useState({ checking: false, signedIn: false, user: null });

  const apiConfigured = isValidHttpUrl(API_URL);

  async function checkAuthSession() {
    setAuthState((s) => ({ ...s, checking: true }));
    try {
      const res = await fetch(`${API_ORIGIN}/api/session`, { credentials: "include" });
      const data = await res.json().catch(() => null);
      const signedIn = Boolean(data?.signedIn);
      const userName = data?.user || null;
      setAuthState({ checking: false, signedIn, user: userName });
    } catch {
      setAuthState({ checking: false, signedIn: false, user: null });
    }
  }

  function signInToFunctionApp() {
    const redirect = encodeURIComponent(window.location.href);
    window.location.href = `${API_ORIGIN}/.auth/login/aad?post_login_redirect_uri=${redirect}`;
  }

  function signOutOfFunctionApp() {
    const redirect = encodeURIComponent(window.location.href);
    window.location.href = `${API_ORIGIN}/.auth/logout?post_logout_redirect_uri=${redirect}`;
  }

  async function onBulkFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const text = await file.text();
      const { entries, errors } = parseEntraTemplateCsv(text, { resetRedemption: resetRedemptionFlag });
      setBulkEntriesFromFile(entries);
      setBulkFileName(file.name);
      setResult({
        ok: errors.length === 0,
        data: {
          parsedFromFile: file.name,
          validRows: entries.length,
          invalidRows: errors,
        },
      });
    } catch {
      setBulkEntriesFromFile([]);
      setBulkFileName(file.name);
      setResult({ ok: false, data: { error: "Unable to read the uploaded CSV." } });
    }
  }

  async function inviteGuest(event) {
    event.preventDefault();

    if (!apiConfigured) {
      setResult({ ok: false, data: { error: "Enter a valid API URL to continue." } });
      return;
    }

    setLoading(true);
    setResult(null);

    if (mode === "single") {
      const normalizedEmail = normalizeEmail(email);
      const normalizedFirstName = sanitizeName(firstName);
      const normalizedLastName = sanitizeName(lastName);
      const normalizedRedirect = redirectUrl.trim();

      if (!EMAIL_REGEX.test(normalizedEmail)) {
        setLoading(false);
        setResult({ ok: false, data: { error: "Please enter a valid guest email." } });
        return;
      }

      if (!isValidHttpUrl(normalizedRedirect)) {
        setLoading(false);
        setResult({ ok: false, data: { error: "Please enter a valid redirect URL." } });
        return;
      }

      const response = await sendInvite(API_URL, {
        email: normalizedEmail,
        firstName: normalizedFirstName,
        lastName: normalizedLastName,
        displayName: getDisplayName(normalizedFirstName, normalizedLastName),
        inviteRedirectUrl: normalizedRedirect,
        sendEmail: sendEmailFlag,
        customizedMessageBody: customizedMessageBody.trim(),
        resetRedemption: resetRedemptionFlag,
      });

      setResult(response);

      if (response.ok) {
        setEmail("");
        setFirstName("");
        setLastName("");
      }

      setLoading(false);
      return;
    }

    let entries = [];
    let validationErrors = [];

    if (bulkEntriesFromFile.length > 0) {
      entries = bulkEntriesFromFile.map((entry) => ({
        ...entry,
        resetRedemption: resetRedemptionFlag,
      }));
    } else {
      if (!isValidHttpUrl(redirectUrl.trim())) {
        setLoading(false);
        setResult({ ok: false, data: { error: "Please enter a valid redirect URL for bulk invites." } });
        return;
      }

      const parsed = parseManualBulk(bulkInput, {
        redirectUrl: redirectUrl.trim(),
        sendEmail: sendEmailFlag,
        customizedMessageBody: customizedMessageBody.trim(),
        resetRedemption: resetRedemptionFlag,
      });
      entries = parsed.entries;
      validationErrors = parsed.errors;
    }

    if (!entries.length) {
      setLoading(false);
      setResult({
        ok: false,
        data: { error: "No valid rows found. Upload Entra CSV or add manual rows." },
      });
      return;
    }

    if (entries.length > MAX_BULK_INVITES) {
      setLoading(false);
      setResult({
        ok: false,
        data: { error: `Bulk invite limit is ${MAX_BULK_INVITES} rows per submission.` },
      });
      return;
    }

    const failures = [];

    for (const entry of entries) {
      const response = await sendInvite(API_URL, entry);
      if (!response.ok) {
        failures.push({ email: entry.email, error: response.data?.error || "Invite failed" });
      }
    }

    const summary = {
      source: bulkEntriesFromFile.length > 0 ? `file:${bulkFileName}` : "manual",
      totalRows: entries.length + validationErrors.length,
      validRows: entries.length,
      invalidRows: validationErrors,
      invited: entries.length - failures.length,
      failed: failures.length,
      failures,
    };

    setResult({ ok: failures.length === 0 && validationErrors.length === 0, data: summary });

    if (failures.length === 0 && validationErrors.length === 0) {
      setBulkInput("");
      setBulkEntriesFromFile([]);
      setBulkFileName("");
    }

    setLoading(false);
  }

  return (
    <div className="app-shell">
      <div className="app-card">
        <header className="app-header">
          <h1>Guest Invite Manager</h1>
          <p>Invite users individually or in bulk using the Entra template.</p>
        </header>

        <div className="auth-bar">
          <span>
            {authState.signedIn
              ? `Signed in${authState.user ? ` as ${authState.user}` : ""}`
              : "Not signed in to Function App"}
          </span>
          <div className="auth-actions">
            <button type="button" className="mode-pill" onClick={checkAuthSession} disabled={authState.checking}>
              {authState.checking ? "Checking..." : "Check Session"}
            </button>
            {!authState.signedIn ? (
              <button type="button" className="mode-pill active" onClick={signInToFunctionApp}>
                Sign In
              </button>
            ) : (
              <button type="button" className="mode-pill" onClick={signOutOfFunctionApp}>
                Sign Out
              </button>
            )}
          </div>
        </div>

        {!apiConfigured && <div className="status status-warning">Invite API URL is invalid in code.</div>}
        {apiConfigured && !authState.signedIn && (
          <div className="status status-warning">
            Sign in first. Backend is restricted to tenant member users.
          </div>
        )}

        <form onSubmit={inviteGuest} className="invite-form">
          <div className="mode-switch" role="tablist" aria-label="Invite mode">
            <button
              type="button"
              className={mode === "single" ? "mode-pill active" : "mode-pill"}
              onClick={() => {
                setMode("single");
                setResult(null);
              }}
            >
              Single Invite
            </button>
            <button
              type="button"
              className={mode === "bulk" ? "mode-pill active" : "mode-pill"}
              onClick={() => {
                setMode("bulk");
                setResult(null);
              }}
            >
              Bulk Invite
            </button>
          </div>

          <div className="field-grid">
            <div className="field-group field-span-2">
              <label htmlFor="redirectUrl">Redirect URL</label>
              <input
                id="redirectUrl"
                type="url"
                value={redirectUrl}
                onChange={(e) => setRedirectUrl(e.target.value)}
                placeholder="https://myapplications.microsoft.com"
                required
              />
            </div>

            <div className="field-group field-span-2 checkbox-row">
              <label>
                <input
                  type="checkbox"
                  checked={sendEmailFlag}
                  onChange={(e) => setSendEmailFlag(e.target.checked)}
                />
                Send invitation email
              </label>
            </div>

            <div className="field-group field-span-2 checkbox-row">
              <label>
                <input
                  type="checkbox"
                  checked={resetRedemptionFlag}
                  onChange={(e) => setResetRedemptionFlag(e.target.checked)}
                />
                Resend invite (reset redemption for existing pending guests)
              </label>
            </div>

            <div className="field-group field-span-2">
              <label htmlFor="customMsg">Custom invitation message (optional)</label>
              <textarea
                id="customMsg"
                rows={3}
                value={customizedMessageBody}
                onChange={(e) => setCustomizedMessageBody(e.target.value)}
                placeholder="Welcome to our organization"
              />
            </div>

            {mode === "single" ? (
              <>
                <div className="field-group field-span-2">
                  <label htmlFor="email">Guest email</label>
                  <input
                    id="email"
                    type="email"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    placeholder="guest@external.com"
                    required
                  />
                </div>

                <div className="field-group">
                  <label htmlFor="firstName">First name (optional)</label>
                  <input
                    id="firstName"
                    value={firstName}
                    onChange={(e) => setFirstName(e.target.value)}
                    maxLength={MAX_NAME_LENGTH}
                    placeholder="Guest"
                  />
                </div>

                <div className="field-group">
                  <label htmlFor="lastName">Last name (optional)</label>
                  <input
                    id="lastName"
                    value={lastName}
                    onChange={(e) => setLastName(e.target.value)}
                    maxLength={MAX_NAME_LENGTH}
                    placeholder="User"
                  />
                </div>
              </>
            ) : (
              <>
                <div className="field-group field-span-2">
                  <label htmlFor="bulkFile">Upload Entra CSV template</label>
                  <input id="bulkFile" type="file" accept=".csv,text/csv" onChange={onBulkFileChange} />
                  {bulkFileName && <small>Loaded file: {bulkFileName}</small>}
                </div>

                <div className="field-group field-span-2">
                  <label htmlFor="bulkInput">Manual bulk rows (fallback)</label>
                  <textarea
                    id="bulkInput"
                    rows={6}
                    value={bulkInput}
                    onChange={(e) => {
                      setBulkInput(e.target.value);
                      if (bulkEntriesFromFile.length) {
                        setBulkEntriesFromFile([]);
                        setBulkFileName("");
                      }
                    }}
                    placeholder={"user1@external.com, First, Last\\nuser2@external.com, Another, User"}
                  />
                  <small>Format: email, first name, last name</small>
                </div>
              </>
            )}
          </div>

          <button type="submit" className="submit-btn" disabled={loading || !apiConfigured || !authState.signedIn}>
            {loading ? "Processing..." : mode === "single" ? "Send Invite" : "Run Bulk Invite"}
          </button>
        </form>

        {result && (
          <section className={result.ok ? "status status-success" : "status status-error"}>
            <strong>{result.ok ? "Completed" : "Completed with issues"}</strong>
            {getRedeemUrl(result) && (
              <p>
                Redemption URL:{" "}
                <a href={getRedeemUrl(result)} target="_blank" rel="noreferrer">
                  Open invite link
                </a>
              </p>
            )}
            <pre>{JSON.stringify(result.data, null, 2)}</pre>
          </section>
        )}
      </div>
    </div>
  );
}
