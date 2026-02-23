const { app } = require("@azure/functions");
const { DefaultAzureCredential, ManagedIdentityCredential } = require("@azure/identity");
const fetch = require("node-fetch");

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const EXT_FIELD = "extension_9722db236acf4a89b1c3463d9a982b12_guestSource";
const GUEST_SOURCE_VALUE = process.env.GUEST_SOURCE_VALUE || "GuestInvited";
const WHOAMI_ENABLED = process.env.ENABLE_WHOAMI === "true" || !process.env.WEBSITE_INSTANCE_ID;
const IS_AZURE_RUNTIME = Boolean(process.env.WEBSITE_INSTANCE_ID);
const ALLOW_LOCAL_ANONYMOUS_INVITE = process.env.ALLOW_LOCAL_ANONYMOUS_INVITE !== "false";
const ALLOWED_INVITER_GROUP_OBJECT_IDS = (process.env.ALLOWED_INVITER_GROUP_OBJECT_IDS || "")
  .split(",")
  .map((v) => v.trim())
  .filter(Boolean);
const ENFORCE_GROUP_MEMBERSHIP_IN_AZURE = process.env.ENFORCE_GROUP_MEMBERSHIP_IN_AZURE !== "false";

function tryParseJson(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function decodeJwtClaims(token) {
  if (!token || typeof token !== "string") return null;
  const parts = token.split(".");
  if (parts.length < 2) return null;

  try {
    const normalized = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const payload = Buffer.from(padded, "base64").toString("utf8");
    return JSON.parse(payload);
  } catch {
    return null;
  }
}

function getHeader(request, key) {
  return request.headers?.get?.(key) || null;
}

function decodeClientPrincipal(request) {
  const raw = getHeader(request, "x-ms-client-principal");
  if (!raw) return null;

  try {
    return JSON.parse(Buffer.from(raw, "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function getClaimValue(principal, claimName) {
  const claims = Array.isArray(principal?.claims) ? principal.claims : [];
  const match = claims.find((c) => c?.typ === claimName);
  return match?.val || null;
}

function getCallerObjectId(principal) {
  return (
    getClaimValue(principal, "http://schemas.microsoft.com/identity/claims/objectidentifier") ||
    getClaimValue(principal, "oid") ||
    principal?.userId ||
    null
  );
}

function getRuntimeCredential() {
  const configuredClientId = (process.env.AZURE_CLIENT_ID || "").trim();

  if (IS_AZURE_RUNTIME) {
    if (configuredClientId) {
      return {
        credential: new ManagedIdentityCredential(configuredClientId),
        mode: "managed_identity_user_assigned",
        azClientIdConfigured: true,
      };
    }

    return {
      credential: new ManagedIdentityCredential(),
      mode: "managed_identity_system_assigned",
      azClientIdConfigured: false,
    };
  }

  return {
    credential: new DefaultAzureCredential(),
    mode: "local_default_credential",
    azClientIdConfigured: Boolean(configuredClientId),
  };
}

async function getGraphAccessToken() {
  const runtime = getRuntimeCredential();
  const token = await runtime.credential.getToken(GRAPH_SCOPE);
  return { token: token?.token, runtime };
}

async function graphRequest(method, url, token, body) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  const text = await res.text();
  const json = tryParseJson(text);

  if (!res.ok) {
    const err = new Error(`Graph call failed: ${res.status} ${res.statusText}`);
    err.status = res.status;
    err.details = json;
    err.wwwAuthenticate = res.headers.get("www-authenticate") || null;
    err.requestId = res.headers.get("request-id") || null;
    err.clientRequestId = res.headers.get("client-request-id") || null;
    throw err;
  }

  return json;
}

async function authorizeInviter(request, graphToken) {
  if (!IS_AZURE_RUNTIME && ALLOW_LOCAL_ANONYMOUS_INVITE) {
    return {
      authorized: true,
      mode: "local_bypass",
      caller: { userType: "Member", id: "local-dev" },
    };
  }

  const principal = decodeClientPrincipal(request);
  if (!principal) {
    return {
      authorized: false,
      status: 401,
      error:
        "Unauthorized. Enable App Service Authentication and call this API as a signed-in internal user.",
    };
  }

  const callerOid = getCallerObjectId(principal);
  if (!callerOid) {
    return {
      authorized: false,
      status: 401,
      error: "Unauthorized. Caller object ID claim is missing.",
    };
  }

  const caller = await graphRequest(
    "GET",
    `${GRAPH_BASE}/users/${callerOid}?$select=id,displayName,userPrincipalName,userType,accountEnabled`,
    graphToken
  );

  if (!caller?.id || caller.accountEnabled === false) {
    return {
      authorized: false,
      status: 403,
      error: "Forbidden. Caller account is not active.",
    };
  }

  if (caller.userType !== "Member") {
    return {
      authorized: false,
      status: 403,
      error: "Forbidden. Only tenant Member users can send invites.",
    };
  }

  if (IS_AZURE_RUNTIME && ENFORCE_GROUP_MEMBERSHIP_IN_AZURE && ALLOWED_INVITER_GROUP_OBJECT_IDS.length === 0) {
    return {
      authorized: false,
      status: 500,
      error: "Server misconfiguration. ALLOWED_INVITER_GROUP_OBJECT_IDS must be set in Azure.",
    };
  }

  const shouldCheckGroupMembership =
    ALLOWED_INVITER_GROUP_OBJECT_IDS.length > 0 &&
    (!IS_AZURE_RUNTIME || ENFORCE_GROUP_MEMBERSHIP_IN_AZURE);

  if (shouldCheckGroupMembership) {
    const membership = await graphRequest(
      "POST",
      `${GRAPH_BASE}/users/${callerOid}/checkMemberGroups`,
      graphToken,
      { groupIds: ALLOWED_INVITER_GROUP_OBJECT_IDS }
    );

    const matched = Array.isArray(membership?.value) ? membership.value : [];
    if (matched.length === 0) {
      return {
        authorized: false,
        status: 403,
        error: "Forbidden. Caller is not in the allowed inviter group.",
      };
    }
  }

  return {
    authorized: true,
    mode: "entra_authenticated",
    caller,
  };
}

app.http("whoami", {
  methods: ["GET"],
  authLevel: "function",
  handler: async (_request, context) => {
    if (!WHOAMI_ENABLED) {
      return { status: 404, jsonBody: { error: "Not found" } };
    }

    try {
      const { token: accessToken, runtime } = await getGraphAccessToken();
      const claims = decodeJwtClaims(accessToken);
      const org = await graphRequest(
        "GET",
        `${GRAPH_BASE}/organization?$select=id,displayName,verifiedDomains`,
        accessToken
      );

      return {
        status: 200,
        jsonBody: {
          runtime: runtime.mode,
          azClientIdConfigured: runtime.azClientIdConfigured,
          tokenClaims: claims
            ? {
                aud: claims.aud,
                appid: claims.appid,
                oid: claims.oid,
                tid: claims.tid,
                roles: claims.roles || [],
                iss: claims.iss,
              }
            : null,
          organization: org?.value?.[0] || null,
        },
      };
    } catch (e) {
      context.error("whoami failed", e);
      return {
        status: e.status || 500,
        jsonBody: {
          error: e.message || "Unknown error",
          details: e.details || null,
          requestId: e.requestId || null,
          clientRequestId: e.clientRequestId || null,
          wwwAuthenticate: e.wwwAuthenticate || null,
        },
      };
    }
  },
});

app.http("inviteGuest", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    let runtime = null;

    try {
      const body = await request.json().catch(() => ({}));
      const email = body.email;
      const displayName =
        body.displayName ||
        [body.firstName, body.lastName].filter(Boolean).join(" ").trim() ||
        email;
      const redirectUrl =
        body.redirectUrl ||
        body.inviteRedirectUrl ||
        body.inviteRedirectURL ||
        "https://myapps.microsoft.com";
      const sendInvitationMessage =
        typeof body.sendEmail === "boolean" ? body.sendEmail : true;
      const customizedMessageBody =
        typeof body.customizedMessageBody === "string"
          ? body.customizedMessageBody.trim()
          : "";
      const messageLanguage =
        typeof body.messageLanguage === "string" && body.messageLanguage.trim()
          ? body.messageLanguage.trim()
          : "en-US";
      const resetRedemption =
        typeof body.resetRedemption === "boolean" ? body.resetRedemption : false;

      if (!email) {
        return { status: 400, jsonBody: { error: "email is required" } };
      }

      const tokenResponse = await getGraphAccessToken();
      const accessToken = tokenResponse.token;
      runtime = tokenResponse.runtime;
      const claims = decodeJwtClaims(accessToken);

      const inviterAuth = await authorizeInviter(request, accessToken);
      if (!inviterAuth.authorized) {
        return {
          status: inviterAuth.status,
          jsonBody: {
            error: inviterAuth.error,
            runtime: runtime?.mode || null,
            azClientIdConfigured: runtime?.azClientIdConfigured || false,
          },
        };
      }

      const invitationBody = {
        invitedUserEmailAddress: email,
        invitedUserDisplayName: displayName,
        inviteRedirectUrl: redirectUrl,
        sendInvitationMessage,
      };

      if (sendInvitationMessage) {
        invitationBody.invitedUserMessageInfo = {
          messageLanguage,
          ...(customizedMessageBody ? { customizedMessageBody } : {}),
        };
      }

      if (resetRedemption) {
        invitationBody.resetRedemption = true;
      }

      const invitation = await graphRequest(
        "POST",
        `${GRAPH_BASE}/invitations`,
        accessToken,
        invitationBody
      );

      const invitedUserId = invitation?.invitedUser?.id;
      if (!invitedUserId) {
        return {
          status: 500,
          jsonBody: {
            error: "Invite succeeded but invitedUser.id not returned",
            invitation,
          },
        };
      }

      await graphRequest(
        "PATCH",
        `${GRAPH_BASE}/users/${invitedUserId}`,
        accessToken,
        { [EXT_FIELD]: GUEST_SOURCE_VALUE }
      );

      return {
        status: 200,
        jsonBody: {
          status: "invited_and_stamped",
          invitedUserId,
          email,
          guestSource: GUEST_SOURCE_VALUE,
          inviteRedeemUrl: invitation?.inviteRedeemUrl || null,
          inviterAuthorizationMode: inviterAuth.mode,
          inviter: {
            id: inviterAuth.caller?.id || null,
            userType: inviterAuth.caller?.userType || null,
            userPrincipalName: inviterAuth.caller?.userPrincipalName || null,
          },
          runtime: runtime?.mode || null,
          azClientIdConfigured: runtime?.azClientIdConfigured || false,
          tokenContext: claims
            ? {
                aud: claims.aud,
                tid: claims.tid,
                appid: claims.appid,
                oid: claims.oid,
                roles: claims.roles || [],
              }
            : null,
        },
      };
    } catch (e) {
      context.error("InviteGuest failed", e);
      return {
        status: e.status || 500,
        jsonBody: {
          error: e.message || "Unknown error",
          details: e.details || null,
          requestId: e.requestId || null,
          clientRequestId: e.clientRequestId || null,
          wwwAuthenticate: e.wwwAuthenticate || null,
          runtime: runtime?.mode || null,
          azClientIdConfigured: runtime?.azClientIdConfigured || false,
        },
      };
    }
  },
});
