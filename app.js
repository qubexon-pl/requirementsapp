(() => {
  const config = window.APP_CONFIG;
  const hasValidConfig = Boolean(config?.msal?.clientId && config.msal.clientId !== 'YOUR-ENTRA-APP-CLIENT-ID');

  const els = {
    signInBtn: document.getElementById('signInBtn'),
    signOutBtn: document.getElementById('signOutBtn'),
    resetConversationBtn: document.getElementById('resetConversationBtn'),
    clearLogsBtn: document.getElementById('clearLogsBtn'),
    authBadge: document.getElementById('authBadge'),
    authDescription: document.getElementById('authDescription'),
    userDisplay: document.getElementById('userDisplay'),
    conversationDisplay: document.getElementById('conversationDisplay'),
    statusText: document.getElementById('statusText'),
    messageList: document.getElementById('messageList'),
    chatForm: document.getElementById('chatForm'),
    messageInput: document.getElementById('messageInput'),
    sendBtn: document.getElementById('sendBtn'),
    logOutput: document.getElementById('logOutput'),
    agentUriDisplay: document.getElementById('agentUriDisplay'),
    chatUriDisplay: document.getElementById('chatUriDisplay'),
    tenantDisplay: document.getElementById('tenantDisplay'),
    clientDisplay: document.getElementById('clientDisplay'),
    scopesDisplay: document.getElementById('scopesDisplay')
  };

  const state = {
    account: null,
    accessToken: null,
    conversationId: null,
    msalClient: null
  };

  const authority = `https://login.microsoftonline.com/${config.msal.tenantId}`;
  const request = { scopes: config.powerPlatform.scopes };

  function log(message, payload) {
    const stamp = new Date().toISOString();
    const line = payload ? `${stamp} ${message}\n${JSON.stringify(payload, null, 2)}` : `${stamp} ${message}`;
    els.logOutput.textContent = `${line}\n\n${els.logOutput.textContent}`.trim();
  }

  function setStatus(text) {
    els.statusText.textContent = text;
  }

  function updateConfigSummary() {
    els.agentUriDisplay.textContent = config.powerPlatform.agentUri;
    els.chatUriDisplay.textContent = config.powerPlatform.agentUri;
    els.tenantDisplay.textContent = config.msal.tenantId;
    els.clientDisplay.textContent = config.msal.clientId;
    els.scopesDisplay.textContent = config.powerPlatform.scopes.join(', ');
  }

  function setAuthUi(account) {
    const signedIn = Boolean(account);
    els.signInBtn.disabled = signedIn;
    els.signOutBtn.disabled = !signedIn;
    els.sendBtn.disabled = !signedIn;
    els.authBadge.textContent = signedIn ? 'Signed in' : 'Signed out';
    els.authBadge.className = `badge ${signedIn ? 'success' : 'muted'}`;
    els.authDescription.textContent = signedIn
      ? 'Authenticated with Microsoft Entra ID. Ready to call the Power Platform agent.'
      : 'Sign in with Entra ID to start a secure conversation.';
    els.userDisplay.textContent = signedIn
      ? `${account.name || 'Unknown user'} (${account.username})`
      : 'Not signed in';
  }

  function addMessage(role, text, meta = role === 'user' ? 'You' : 'Agent') {
    const article = document.createElement('article');
    article.className = `message ${role}`;

    const header = document.createElement('div');
    header.className = 'message-meta';
    header.textContent = meta;

    const body = document.createElement('div');
    body.className = 'message-body';
    body.textContent = text;

    article.append(header, body);
    els.messageList.appendChild(article);
    els.messageList.scrollTop = els.messageList.scrollHeight;
  }

  function setConversation(id) {
    state.conversationId = id;
    els.conversationDisplay.textContent = id || 'Not started';
  }

  async function ensureConversation() {
    if (state.conversationId) return state.conversationId;

    setStatus('Starting conversation...');
    log('Creating a new conversation');
    const response = await fetch(config.powerPlatform.agentUri, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ locale: config.chat.locale })
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      log('Conversation creation failed', data);
      throw new Error(data?.error?.message || data?.message || `Conversation creation failed with status ${response.status}`);
    }

    const conversationId = data.conversationId || data.id || data?.conversation?.id;
    if (!conversationId) {
      log('Conversation creation response did not contain a conversation id', data);
      throw new Error('The agent API responded without a conversation id. Check the endpoint and payload requirements.');
    }

    setConversation(conversationId);
    log('Conversation created', data);
    addMessage('system', `Conversation started: ${conversationId}`, 'System');
    return conversationId;
  }

  async function acquireToken() {
    if (!state.account) throw new Error('Sign in first.');

    try {
      const result = await state.msalClient.acquireTokenSilent({ ...request, account: state.account });
      state.accessToken = result.accessToken;
      log('Access token acquired silently');
      return result.accessToken;
    } catch (error) {
      log('Silent token acquisition failed, falling back to popup', { message: error.message });
      const result = await state.msalClient.acquireTokenPopup(request);
      state.accessToken = result.accessToken;
      return result.accessToken;
    }
  }

  async function sendMessage(messageText) {
    await acquireToken();
    const conversationId = await ensureConversation();
    const activitiesUri = `${config.powerPlatform.agentUri.replace(/\?.*$/, '')}/${encodeURIComponent(conversationId)}/activities?api-version=2022-03-01-preview`;

    const payload = {
      type: 'message',
      text: messageText,
      locale: config.chat.locale,
      from: {
        id: state.account?.homeAccountId || state.account?.username || 'web-user',
        name: state.account?.name || state.account?.username || 'Web user'
      }
    };

    setStatus('Sending message...');
    log('Sending activity', { activitiesUri, payload });

    const response = await fetch(activitiesUri, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      log('Message send failed', data);
      throw new Error(data?.error?.message || data?.message || `Message send failed with status ${response.status}`);
    }

    log('Activity response received', data);
    const replies = extractReplies(data);
    if (replies.length === 0) {
      addMessage('system', 'The agent accepted the request, but no text reply was returned in the immediate response. Inspect Diagnostics for the raw payload.', 'System');
    } else {
      replies.forEach((reply) => addMessage('agent', reply));
    }

    setStatus('Idle');
  }

  function extractReplies(data) {
    const collected = [];
    const addIfText = (value) => {
      if (typeof value === 'string' && value.trim()) collected.push(value.trim());
    };

    if (Array.isArray(data?.activities)) {
      data.activities.forEach((activity) => {
        addIfText(activity?.text);
        if (Array.isArray(activity?.attachments)) {
          activity.attachments.forEach((attachment) => addIfText(attachment?.content?.text));
        }
      });
    }

    addIfText(data?.text);
    addIfText(data?.message);
    addIfText(data?.reply);
    addIfText(data?.output);

    if (Array.isArray(data?.messages)) {
      data.messages.forEach((message) => addIfText(message?.text || message?.content));
    }

    return [...new Set(collected)];
  }

  async function signIn() {
    setStatus('Signing in...');
    const result = await state.msalClient.loginPopup(request);
    state.account = result.account;
    state.msalClient.setActiveAccount(state.account);
    setAuthUi(state.account);
    await acquireToken();
    addMessage('system', `${state.account.name || state.account.username} signed in successfully.`, 'System');
    setStatus('Idle');
  }

  async function signOut() {
    setStatus('Signing out...');
    setConversation(null);
    state.accessToken = null;
    await state.msalClient.logoutPopup({
      account: state.account,
      postLogoutRedirectUri: config.msal.postLogoutRedirectUri || window.location.origin
    });
    state.account = null;
    setAuthUi(null);
    addMessage('system', 'Signed out.', 'System');
    setStatus('Idle');
  }

  async function initialize() {
    updateConfigSummary();
    if (!hasValidConfig) {
      setAuthUi(null);
      els.signInBtn.disabled = true;
      els.sendBtn.disabled = true;
      addMessage('system', 'Update config.js with your Entra ID tenant and client ID before signing in.', 'System');
      log('Configuration incomplete. Update config.js before testing authentication.');
      setStatus('Configuration required');
      return;
    }

    state.msalClient = new msal.PublicClientApplication({
      auth: {
        clientId: config.msal.clientId,
        authority,
        redirectUri: config.msal.redirectUri,
        postLogoutRedirectUri: config.msal.postLogoutRedirectUri || config.msal.redirectUri
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
      }
    });

    await state.msalClient.initialize();
    const account = state.msalClient.getActiveAccount() || state.msalClient.getAllAccounts()[0] || null;
    if (account) {
      state.account = account;
      state.msalClient.setActiveAccount(account);
      setAuthUi(account);
      try {
        await acquireToken();
      } catch (error) {
        log('Initial token acquisition failed', { message: error.message });
      }
    } else {
      setAuthUi(null);
      els.sendBtn.disabled = true;
    }

    els.signInBtn.addEventListener('click', () => signIn().catch(handleError));
    els.signOutBtn.addEventListener('click', () => signOut().catch(handleError));
    els.resetConversationBtn.addEventListener('click', () => {
      setConversation(null);
      addMessage('system', 'Conversation state reset. The next message will create a new conversation.', 'System');
      log('Conversation reset by user');
    });
    els.clearLogsBtn.addEventListener('click', () => {
      els.logOutput.textContent = 'Logs cleared.';
    });

    els.chatForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      const text = els.messageInput.value.trim();
      if (!text) return;
      addMessage('user', text);
      els.messageInput.value = '';
      try {
        await sendMessage(text);
      } catch (error) {
        handleError(error);
      }
    });

    setStatus('Idle');
  }

  function handleError(error) {
    console.error(error);
    const message = error?.message || 'Unexpected error';
    addMessage('error', message, 'Error');
    log('Error', { message });
    setStatus('Error');
  }

  initialize().catch(handleError);
})();
