/**
 * OAuth Utility Functions for Cloudflare Workers OAuth Provider
 *
 * Handles client approval flow and cookie-based state management for
 * OAuth 2.1 + PKCE authorization with Microsoft Identity Platform.
 *
 * ARCHITECTURE:
 * - Uses encrypted cookies for CSRF protection and state preservation
 * - Implements HTML-based approval dialog for OAuth consent
 * - Provides utilities for parsing OAuth redirect responses
 *
 * SECURITY DESIGN:
 * - HMAC signing prevents cookie tampering
 * - HttpOnly flag prevents XSS access
 * - SameSite=Lax prevents CSRF attacks
 * - 15-minute expiry limits exposure window
 * - All sensitive data sanitized before rendering
 *
 * TEMPLATE ARCHITECTURE:
 * Inline CSS/HTML chosen over external assets for:
 * - Zero external dependencies (security)
 * - Consistent rendering across clients
 * - Cloudflare Workers edge compatibility
 * Consider extracting to external template if maintenance becomes complex.
 */

import type {
  AuthRequest,
  ClientInfo,
} from "@cloudflare/workers-oauth-provider";

const COOKIE_NAME = "mcp-approved-clients";
const ONE_YEAR_IN_SECONDS = 31536000;

// ============================================================================
// COOKIE CONFIGURATION
// ============================================================================

/**
 * Encodes arbitrary data to a URL-safe base64 string.
 * @param data - The data to encode (will be stringified).
 * @returns A URL-safe base64 encoded string.
 */
function _encodeState(data: any): string {
  try {
    const jsonString = JSON.stringify(data);
    /** Base64 encoding for URL-safe state transmission */
    return btoa(jsonString);
  } catch (e) {
    throw new Error("Could not encode state");
  }
}

/**
 * Decodes a URL-safe base64 string back to its original data.
 * @param encoded - The URL-safe base64 encoded string.
 * @returns The original data.
 */
function decodeState<T = any>(encoded: string): T {
  try {
    const jsonString = atob(encoded);
    return JSON.parse(jsonString);
  } catch (e) {
    throw new Error("Could not decode state");
  }
}

// ============================================================================
// CRYPTOGRAPHIC UTILITIES
// ============================================================================

/**
 * Imports a secret key string for HMAC-SHA256 signing
 *
 * SECURITY: Creates non-extractable CryptoKey to prevent key material exposure
 * in JavaScript context. Key is only accessible via sign/verify operations.
 *
 * @param secret - Raw secret key string (should be cryptographically random)
 * @returns Promise resolving to CryptoKey for HMAC operations
 * @throws Error if secret is undefined or empty
 * @private
 */
async function importKey(secret: string): Promise<CryptoKey> {
  if (!secret) {
    throw new Error(
      "COOKIE_SECRET is not defined. A secret key is required for signing cookies.",
    );
  }
  const enc = new TextEncoder();
  return crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { hash: "SHA-256", name: "HMAC" },
    false /** Non-extractable for security */,
    ["sign", "verify"] /** Required operations for HMAC */,
  );
}

/**
 * Signs data using HMAC-SHA256.
 * @param key - The CryptoKey for signing.
 * @param data - The string data to sign.
 * @returns A promise resolving to the signature as a hex string.
 */
async function signData(key: CryptoKey, data: string): Promise<string> {
  const enc = new TextEncoder();
  const signatureBuffer = await crypto.subtle.sign(
    "HMAC",
    key,
    enc.encode(data),
  );
  /** Convert ArrayBuffer to hex string for cookie storage */
  return Array.from(new Uint8Array(signatureBuffer))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

/**
 * Verifies an HMAC-SHA256 signature.
 * @param key - The CryptoKey for verification.
 * @param signatureHex - The signature to verify (hex string).
 * @param data - The original data that was signed.
 * @returns A promise resolving to true if the signature is valid, false otherwise.
 */
async function verifySignature(
  key: CryptoKey,
  signatureHex: string,
  data: string,
): Promise<boolean> {
  const enc = new TextEncoder();
  try {
    /** Convert hex signature back to ArrayBuffer for verification */
    const signatureBytes = new Uint8Array(
      signatureHex.match(/.{1,2}/g)!.map((byte) => Number.parseInt(byte, 16)),
    );
    return await crypto.subtle.verify(
      "HMAC",
      key,
      signatureBytes.buffer,
      enc.encode(data),
    );
  } catch (e) {
    /** Invalid signature format or verification failure */
    return false;
  }
}

/**
 * Parses the signed cookie and verifies its integrity.
 * @param cookieHeader - The value of the Cookie header from the request.
 * @param secret - The secret key used for signing.
 * @returns A promise resolving to the list of approved client IDs if the cookie is valid, otherwise null.
 */
async function getApprovedClientsFromCookie(
  cookieHeader: string | null,
  secret: string,
): Promise<string[] | null> {
  if (!cookieHeader) return null;

  const cookies = cookieHeader.split(";").map((c) => c.trim());
  const targetCookie = cookies.find((c) => c.startsWith(`${COOKIE_NAME}=`));

  if (!targetCookie) return null;

  const cookieValue = targetCookie.substring(COOKIE_NAME.length + 1);
  const parts = cookieValue.split(".");

  if (parts.length !== 2) {
    /** Invalid cookie format - expected signature.payload */
    return null;
  }

  const [signatureHex, base64Payload] = parts;
  const payload =
    atob(base64Payload); /** Decode base64 payload to JSON string */

  const key = await importKey(secret);
  const isValid = await verifySignature(key, signatureHex, payload);

  if (!isValid) {
    /** Cookie signature verification failed - possible tampering */
    return null;
  }

  try {
    const approvedClients = JSON.parse(payload);
    if (!Array.isArray(approvedClients)) {
      /** Invalid payload structure - expected array of client IDs */
      return null;
    }
    /** Validate all elements are strings (client IDs) */
    if (!approvedClients.every((item) => typeof item === "string")) {
      return null;
    }
    return approvedClients as string[];
  } catch (e) {
    /** JSON parsing failed - corrupted cookie data */
    return null;
  }
}

/** Exported OAuth utility functions */

/**
 * Checks if a given client ID has already been approved by the user,
 * based on a signed cookie.
 *
 * @param request - The incoming Request object to read cookies from.
 * @param clientId - The OAuth client ID to check approval for.
 * @param cookieSecret - The secret key used to sign/verify the approval cookie.
 * @returns A promise resolving to true if the client ID is in the list of approved clients in a valid cookie, false otherwise.
 */
export async function clientIdAlreadyApproved(
  request: Request,
  clientId: string,
  cookieSecret: string,
): Promise<boolean> {
  if (!clientId) return false;
  const cookieHeader = request.headers.get("Cookie");
  const approvedClients = await getApprovedClientsFromCookie(
    cookieHeader,
    cookieSecret,
  );

  return approvedClients?.includes(clientId) ?? false;
}

/**
 * Configuration for the approval dialog
 */
export interface ApprovalDialogOptions {
  /**
   * Client information to display in the approval dialog
   */
  client: ClientInfo | null;
  /**
   * Server information to display in the approval dialog
   */
  server: {
    name: string;
    logo?: string;
    description?: string;
  };
  /**
   * Arbitrary state data to pass through the approval flow
   * Will be encoded in the form and returned when approval is complete
   */
  state: Record<string, any>;
  /**
   * Name of the cookie to use for storing approvals
   * @default "mcp_approved_clients"
   */
  cookieName?: string;
  /**
   * Secret used to sign cookies for verification
   * Can be a string or Uint8Array
   * @default Built-in Uint8Array key
   */
  cookieSecret?: string | Uint8Array;
  /**
   * Cookie domain
   * @default current domain
   */
  cookieDomain?: string;
  /**
   * Cookie path
   * @default "/"
   */
  cookiePath?: string;
  /**
   * Cookie max age in seconds
   * @default 30 days
   */
  cookieMaxAge?: number;
}

/**
 * Renders an approval dialog for OAuth authorization
 * The dialog displays information about the client and server
 * and includes a form to submit approval
 *
 * @param request - The HTTP request
 * @param options - Configuration for the approval dialog
 * @returns A Response containing the HTML approval dialog
 */
export function renderApprovalDialog(
  request: Request,
  options: ApprovalDialogOptions,
): Response {
  const { client, server, state } = options;

  // Encode state for form submission
  const encodedState = btoa(JSON.stringify(state));

  // Sanitize any untrusted content
  const serverName = sanitizeHtml(server.name);
  const clientName = client?.clientName
    ? sanitizeHtml(client.clientName)
    : "Unknown MCP Client";
  const serverDescription = server.description
    ? sanitizeHtml(server.description)
    : "";

  // Safe URLs
  const logoUrl = server.logo ? sanitizeHtml(server.logo) : "";
  const clientUri = client?.clientUri ? sanitizeHtml(client.clientUri) : "";
  const policyUri = client?.policyUri ? sanitizeHtml(client.policyUri) : "";
  const tosUri = client?.tosUri ? sanitizeHtml(client.tosUri) : "";

  // Client contacts
  const contacts =
    client?.contacts && client.contacts.length > 0
      ? sanitizeHtml(client.contacts.join(", "))
      : "";

  // Get redirect URIs
  const redirectUris =
    client?.redirectUris && client.redirectUris.length > 0
      ? client.redirectUris.map((uri) => sanitizeHtml(uri))
      : [];

  // Generate HTML for the approval dialog
  const htmlContent = `
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${clientName} | Authorization Request</title>
        <style>
          /* Modern, responsive styling with system fonts */
          :root {
            --primary-color: #0070f3;
            --error-color: #f44336;
            --border-color: #e5e7eb;
            --text-color: #333;
            --background-color: #fff;
            --card-shadow: 0 8px 36px 8px rgba(0, 0, 0, 0.1);
          }
          
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, 
                         Helvetica, Arial, sans-serif, "Apple Color Emoji", 
                         "Segoe UI Emoji", "Segoe UI Symbol";
            line-height: 1.6;
            color: var(--text-color);
            background-color: #f9fafb;
            margin: 0;
            padding: 0;
          }
          
          .container {
            max-width: 600px;
            margin: 2rem auto;
            padding: 1rem;
          }
          
          .precard {
            padding: 2rem;
            text-align: center;
          }
          
          .card {
            background-color: var(--background-color);
            border-radius: 8px;
            box-shadow: var(--card-shadow);
            padding: 2rem;
          }
          
          .header {
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 1.5rem;
          }
          
          .logo {
            width: 48px;
            height: 48px;
            margin-right: 1rem;
            border-radius: 8px;
            object-fit: contain;
          }
          
          .title {
            margin: 0;
            font-size: 1.3rem;
            font-weight: 400;
          }
          
          .alert {
            margin: 0;
            font-size: 1.5rem;
            font-weight: 400;
            margin: 1rem 0;
            text-align: center;
          }
          
          .description {
            color: #555;
          }
          
          .client-info {
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 1rem 1rem 0.5rem;
            margin-bottom: 1.5rem;
          }
          
          .client-name {
            font-weight: 600;
            font-size: 1.2rem;
            margin: 0 0 0.5rem 0;
          }
          
          .client-detail {
            display: flex;
            margin-bottom: 0.5rem;
            align-items: baseline;
          }
          
          .detail-label {
            font-weight: 500;
            min-width: 120px;
          }
          
          .detail-value {
            font-family: SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
            word-break: break-all;
          }
          
          .detail-value a {
            color: inherit;
            text-decoration: underline;
          }
          
          .detail-value.small {
            font-size: 0.8em;
          }
          
          .external-link-icon {
            font-size: 0.75em;
            margin-left: 0.25rem;
            vertical-align: super;
          }
          
          .actions {
            display: flex;
            justify-content: flex-end;
            gap: 1rem;
            margin-top: 2rem;
          }
          
          .button {
            padding: 0.75rem 1.5rem;
            border-radius: 6px;
            font-weight: 500;
            cursor: pointer;
            border: none;
            font-size: 1rem;
          }
          
          .button-primary {
            background-color: var(--primary-color);
            color: white;
          }
          
          .button-secondary {
            background-color: transparent;
            border: 1px solid var(--border-color);
            color: var(--text-color);
          }
          
          /* Responsive adjustments */
          @media (max-width: 640px) {
            .container {
              margin: 1rem auto;
              padding: 0.5rem;
            }
            
            .card {
              padding: 1.5rem;
            }
            
            .client-detail {
              flex-direction: column;
            }
            
            .detail-label {
              min-width: unset;
              margin-bottom: 0.25rem;
            }
            
            .actions {
              flex-direction: column;
            }
            
            .button {
              width: 100%;
            }
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="precard">
            <div class="header">
              ${logoUrl ? `<img src="${logoUrl}" alt="${serverName} Logo" class="logo">` : ""}
            <h1 class="title"><strong>${serverName}</strong></h1>
            </div>
            
            ${serverDescription ? `<p class="description">${serverDescription}</p>` : ""}
          </div>
            
          <div class="card">
            
            <h2 class="alert"><strong>${clientName || "A new MCP Client"}</strong> is requesting access</h1>
            
            <div class="client-info">
              <div class="client-detail">
                <div class="detail-label">Name:</div>
                <div class="detail-value">
                  ${clientName}
                </div>
              </div>
              
              ${
                clientUri
                  ? `
                <div class="client-detail">
                  <div class="detail-label">Website:</div>
                  <div class="detail-value small">
                    <a href="${clientUri}" target="_blank" rel="noopener noreferrer">
                      ${clientUri}
                    </a>
                  </div>
                </div>
              `
                  : ""
              }
              
              ${
                policyUri
                  ? `
                <div class="client-detail">
                  <div class="detail-label">Privacy Policy:</div>
                  <div class="detail-value">
                    <a href="${policyUri}" target="_blank" rel="noopener noreferrer">
                      ${policyUri}
                    </a>
                  </div>
                </div>
              `
                  : ""
              }
              
              ${
                tosUri
                  ? `
                <div class="client-detail">
                  <div class="detail-label">Terms of Service:</div>
                  <div class="detail-value">
                    <a href="${tosUri}" target="_blank" rel="noopener noreferrer">
                      ${tosUri}
                    </a>
                  </div>
                </div>
              `
                  : ""
              }
              
              ${
                redirectUris.length > 0
                  ? `
                <div class="client-detail">
                  <div class="detail-label">Redirect URIs:</div>
                  <div class="detail-value small">
                    ${redirectUris.map((uri) => `<div>${uri}</div>`).join("")}
                  </div>
                </div>
              `
                  : ""
              }
              
              ${
                contacts
                  ? `
                <div class="client-detail">
                  <div class="detail-label">Contact:</div>
                  <div class="detail-value">${contacts}</div>
                </div>
              `
                  : ""
              }
            </div>
            
            <p>This MCP Client is requesting to be authorized on ${serverName}. If you approve, you will be redirected to complete authentication.</p>
            
            <form method="post" action="${new URL(request.url).pathname}">
              <input type="hidden" name="state" value="${encodedState}">
              
              <div class="actions">
                <button type="button" class="button button-secondary" onclick="window.history.back()">Cancel</button>
                <button type="submit" class="button button-primary">Approve</button>
              </div>
            </form>
          </div>
        </div>
      </body>
    </html>
  `;

  return new Response(htmlContent, {
    headers: {
      "Content-Type": "text/html; charset=utf-8",
    },
  });
}

/**
 * Result of parsing the approval form submission.
 */
export interface ParsedApprovalResult {
  /** The original state object passed through the form. */
  state: any;
  /** Headers to set on the redirect response, including the Set-Cookie header. */
  headers: Record<string, string>;
}

/**
 * Parses the form submission from the approval dialog, extracts the state,
 * and generates Set-Cookie headers to mark the client as approved.
 *
 * @param request - The incoming POST Request object containing the form data.
 * @param cookieSecret - The secret key used to sign the approval cookie.
 * @returns A promise resolving to an object containing the parsed state and necessary headers.
 * @throws If the request method is not POST, form data is invalid, or state is missing.
 */
export async function parseRedirectApproval(
  request: Request,
  cookieSecret: string,
): Promise<ParsedApprovalResult> {
  if (request.method !== "POST") {
    throw new Error("Invalid request method. Expected POST.");
  }

  let state: any;
  let clientId: string | undefined;

  try {
    const formData = await request.formData();
    const encodedState = formData.get("state");

    if (typeof encodedState !== "string" || !encodedState) {
      throw new Error("Missing or invalid 'state' in form data.");
    }

    state = decodeState<{ oauthReqInfo?: AuthRequest }>(encodedState); // Decode the state
    clientId = state?.oauthReqInfo?.clientId; // Extract clientId from within the state

    if (!clientId) {
      throw new Error("Could not extract clientId from state object.");
    }
  } catch (e) {
    /** Form processing error - invalid or missing state data */
    throw new Error(
      `Failed to parse approval form: ${e instanceof Error ? e.message : String(e)}`,
    );
  }

  // Get existing approved clients
  const cookieHeader = request.headers.get("Cookie");
  const existingApprovedClients =
    (await getApprovedClientsFromCookie(cookieHeader, cookieSecret)) || [];

  // Add the newly approved client ID (avoid duplicates)
  const updatedApprovedClients = Array.from(
    new Set([...existingApprovedClients, clientId]),
  );

  // Sign the updated list
  const payload = JSON.stringify(updatedApprovedClients);
  const key = await importKey(cookieSecret);
  const signature = await signData(key, payload);
  const newCookieValue = `${signature}.${btoa(payload)}`; // signature.base64(payload)

  // Generate Set-Cookie header
  const headers: Record<string, string> = {
    "Set-Cookie": `${COOKIE_NAME}=${newCookieValue}; HttpOnly; Secure; Path=/; SameSite=Lax; Max-Age=${ONE_YEAR_IN_SECONDS}`,
  };

  return { headers, state };
}

/**
 * Sanitizes HTML content to prevent XSS attacks
 * @param unsafe - The unsafe string that might contain HTML
 * @returns A safe string with HTML special characters escaped
 */
function sanitizeHtml(unsafe: string): string {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
