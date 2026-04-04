
const APP_STORAGE_KEY = "ccm-app-state-v1";
const SCRIPT_URL_KEY = "ccm-script-url";
const THEME_KEY = "ccm-theme";
const CATEGORIES = [
  "Fuel",
  "Grocery",
  "Travel",
  "Shopping",
  "Food & Dining",
  "Bills & Utilities",
  "Medical",
  "Insurance",
  "Education",
  "Entertainment",
  "Groceries",
  "Charges",
  "Payment",
  "Others"
];
const REWARD_RULE_CATEGORIES = [
  "Food & Dining",
  "Shopping",
  "Bills & Utilities",
  "Travel",
  "Medical",
  "Insurance",
  "Payment",
  "Groceries",
  "Charges",
  "Education",
  "Entertainment",
  "Fuel",
  "All",
  "Other Spends"
];
const TRANSACTION_NATURES = ["Expense", "Payment", "Refund", "Bonus"];
const TRANSACTION_TYPES = ["Online", "Offline / POS", "UPI", "Wallet", "Bill Payment", "Fuel", "Insurance", "EMI", "Card-to-Card", "ATM Withdrawal", "Other"];
const PAYMENT_STATUSES = ["Unpaid", "Paid", "Partially Paid"];
const STARTER_CARD_NAMES = [
  "HDFC Regalia", "HDFC Swiggy", "HDFC Tata Neo", "HDFC Marriott Bonvoy", "Axis Atlus", "Axis MyZone",
  "Amex Platinum Travel", "Amex Membership", "Amex Platinum Travel Rruchi", "IDFC First Business",
  "ICICI Amazon Pannkaj", "ICICI Sapphiro", "ICICI Amazon Rruchi", "Kotak Indigo", "Kotak PVR",
  "Kotak 811", "Stancy Ultimate", "RBL Bank Bajaj", "RBL Bank Popcorn", "Indusind Easydiner",
  "SBI BPCL", "SBI Prime Pannkaj", "SBI Elite Rruchi"
];

const state = {
  cards: [],
  transactions: [],
  rewardActivities: [],
  rewardRulesCatalog: [],
  milestoneCatalog: [],
  sharedLimitGroups: [],
  settings: { scriptUrl: localStorage.getItem(SCRIPT_URL_KEY) || "" },
  lastSync: null,
  ui: {
    activeView: "dashboard-view",
    activeCardId: null,
    editingCardId: null,
    editingTransactionId: null,
    editingRewardRuleId: null,
    editingMilestoneId: null,
    editingSharedLimitGroupId: null,
    mobileNavOpen: false
  }
};

const elements = {};
let toastTimer = null;

document.addEventListener("DOMContentLoaded", initApp);

function initApp() {
  cacheElements();
  bindEvents();
  loadLocalState();
  applySavedTheme();
  seedStarterData();
  recalculateAllRewardPoints();
  saveLocalState();
  setDefaultDates();
  populateStaticSelects();
  populateCardDropdowns();
  renderTransactionSuggestions();
  updateRewardRuleScopeUI();
  handleViewportChange();
  renderApp();
}

function cacheElements() {
  elements.viewLinks = Array.from(document.querySelectorAll(".view-link"));
  elements.workspaceViews = Array.from(document.querySelectorAll(".workspace-view"));
  elements.navigatorPanel = document.getElementById("navigatorPanel");
  elements.openTransactionModalButton = document.getElementById("openTransactionModalButton");
  elements.openMobileNavButton = document.getElementById("openMobileNavButton");
  elements.closeMobileNavButton = document.getElementById("closeMobileNavButton");
  elements.mobileNewTransactionButton = document.getElementById("mobileNewTransactionButton");
  elements.mobileActiveViewLabel = document.getElementById("mobileActiveViewLabel");
  elements.openCardPanelButton = document.getElementById("openCardPanelButton");
  elements.themeToggleButton = document.getElementById("themeToggleButton");
  elements.navigatorCardCount = document.getElementById("navigatorCardCount");
  elements.navigatorCardList = document.getElementById("navigatorCardList");
  elements.dashboardCardCount = document.getElementById("dashboardCardCount");
  elements.upcomingDues = document.getElementById("upcomingDues");
  elements.highSpendAlertStrip = document.getElementById("highSpendAlertStrip");
  elements.priorityDueList = document.getElementById("priorityDueList");
  elements.statementOrderList = document.getElementById("statementOrderList");
  elements.utilizationList = document.getElementById("utilizationList");
  elements.bankGroupList = document.getElementById("bankGroupList");
  elements.rewardRuleForm = document.getElementById("rewardRuleForm");
  elements.rewardRuleCard = document.getElementById("rewardRuleCard");
  elements.rewardRuleCategory = document.getElementById("rewardRuleCategory");
  elements.rewardRuleCategorySuggestions = document.getElementById("rewardRuleCategorySuggestions");
  elements.rewardRuleScope = document.getElementById("rewardRuleScope");
  elements.rewardRuleMerchant = document.getElementById("rewardRuleMerchant");
  elements.rewardRuleMerchantSuggestions = document.getElementById("rewardRuleMerchantSuggestions");
  elements.rewardRuleBaseUnit = document.getElementById("rewardRuleBaseUnit");
  elements.rewardRuleRewardPerUnit = document.getElementById("rewardRuleRewardPerUnit");
  elements.rewardRuleMultiplier = document.getElementById("rewardRuleMultiplier");
  elements.rewardRulePriority = document.getElementById("rewardRulePriority");
  elements.cancelRewardRuleEditButton = document.getElementById("cancelRewardRuleEditButton");
  elements.rewardRulesTableBody = document.getElementById("rewardRulesTableBody");
  elements.milestoneForm = document.getElementById("milestoneForm");
  elements.milestoneCard = document.getElementById("milestoneCard");
  elements.milestoneType = document.getElementById("milestoneType");
  elements.milestoneCycleType = document.getElementById("milestoneCycleType");
  elements.milestoneTarget = document.getElementById("milestoneTarget");
  elements.milestoneRewardType = document.getElementById("milestoneRewardType");
  elements.milestoneRewardValue = document.getElementById("milestoneRewardValue");
  elements.milestoneNotes = document.getElementById("milestoneNotes");
  elements.cancelMilestoneEditButton = document.getElementById("cancelMilestoneEditButton");
  elements.milestonesTableBody = document.getElementById("milestonesTableBody");
  elements.transactionSearchMerchant = document.getElementById("transactionSearchMerchant");
  elements.transactionSearchDescription = document.getElementById("transactionSearchDescription");
  elements.transactionSearchCard = document.getElementById("transactionSearchCard");
  elements.transactionSearchCategory = document.getElementById("transactionSearchCategory");
  elements.transactionSearchNature = document.getElementById("transactionSearchNature");
  elements.transactionSearchFromDate = document.getElementById("transactionSearchFromDate");
  elements.transactionSearchToDate = document.getElementById("transactionSearchToDate");
  elements.transactionSearchMinAmount = document.getElementById("transactionSearchMinAmount");
  elements.transactionSearchMaxAmount = document.getElementById("transactionSearchMaxAmount");
  elements.clearTransactionSearchButton = document.getElementById("clearTransactionSearchButton");
  elements.transactionSearchSummary = document.getElementById("transactionSearchSummary");
  elements.transactionTableBody = document.getElementById("transactionTableBody");
  elements.accountViewTitle = document.getElementById("accountViewTitle");
  elements.accountViewSubtitle = document.getElementById("accountViewSubtitle");
  elements.backToDashboardButton = document.getElementById("backToDashboardButton");
  elements.accountAvailableLimit = document.getElementById("accountAvailableLimit");
  elements.accountTotalLimit = document.getElementById("accountTotalLimit");
  elements.accountOutstanding = document.getElementById("accountOutstanding");
  elements.accountUnbilled = document.getElementById("accountUnbilled");
  elements.accountAnniversarySpend = document.getElementById("accountAnniversarySpend");
  elements.accountRewardTotal = document.getElementById("accountRewardTotal");
  elements.accountCycleLabel = document.getElementById("accountCycleLabel");
  elements.accountCycleBilled = document.getElementById("accountCycleBilled");
  elements.accountCyclePaid = document.getElementById("accountCyclePaid");
  elements.accountCyclePending = document.getElementById("accountCyclePending");
  elements.accountMetaSummary = document.getElementById("accountMetaSummary");
  elements.accountSearchMerchant = document.getElementById("accountSearchMerchant");
  elements.accountSearchDescription = document.getElementById("accountSearchDescription");
  elements.clearAccountSearchButton = document.getElementById("clearAccountSearchButton");
  elements.accountSearchSummary = document.getElementById("accountSearchSummary");
  elements.accountTransactionTableBody = document.getElementById("accountTransactionTableBody");
  elements.cardForm = document.getElementById("cardForm");
  elements.cardSearchName = document.getElementById("cardSearchName");
  elements.cardSearchGroup = document.getElementById("cardSearchGroup");
  elements.cardSearchCycle = document.getElementById("cardSearchCycle");
  elements.clearCardSearchButton = document.getElementById("clearCardSearchButton");
  elements.cardsSearchSummary = document.getElementById("cardsSearchSummary");
  elements.cardSummaryTotal = document.getElementById("cardSummaryTotal");
  elements.cardSummaryBanks = document.getElementById("cardSummaryBanks");
  elements.cardSummaryShared = document.getElementById("cardSummaryShared");
  elements.cardSummaryDueSoon = document.getElementById("cardSummaryDueSoon");
  elements.sharedLimitGroupForm = document.getElementById("sharedLimitGroupForm");
  elements.sharedLimitGroupName = document.getElementById("sharedLimitGroupName");
  elements.sharedLimitGroupLimit = document.getElementById("sharedLimitGroupLimit");
  elements.sharedLimitGroupMembers = document.getElementById("sharedLimitGroupMembers");
  elements.sharedLimitGroupNotes = document.getElementById("sharedLimitGroupNotes");
  elements.cancelSharedLimitGroupEditButton = document.getElementById("cancelSharedLimitGroupEditButton");
  elements.sharedLimitGroupsTableBody = document.getElementById("sharedLimitGroupsTableBody");
  elements.cardName = document.getElementById("cardName");
  elements.bankName = document.getElementById("bankName");
  elements.creditLimit = document.getElementById("creditLimit");
  elements.billingCycleDate = document.getElementById("billingCycleDate");
  elements.dueDate = document.getElementById("dueDate");
  elements.statementDate = document.getElementById("statementDate");
  elements.anniversaryDate = document.getElementById("anniversaryDate");
  elements.renewalTarget = document.getElementById("renewalTarget");
  elements.sharedGroup = document.getElementById("sharedGroup");
  elements.preferredCategories = document.getElementById("preferredCategories");
  elements.preferredTransactionTypes = document.getElementById("preferredTransactionTypes");
  elements.supportedTransactionNature = document.getElementById("supportedTransactionNature");
  elements.rewardRules = document.getElementById("rewardRules");
  elements.advancedRewardRules = document.getElementById("advancedRewardRules");
  elements.milestones = document.getElementById("milestones");
  elements.annualBenefits = document.getElementById("annualBenefits");
  elements.cancelCardEditButton = document.getElementById("cancelCardEditButton");
  elements.cardTableBody = document.getElementById("cardTableBody");
  elements.settingsForm = document.getElementById("settingsForm");
  elements.scriptUrl = document.getElementById("scriptUrl");
  elements.syncButton = document.getElementById("syncButton");
  elements.replaceWithRemoteButton = document.getElementById("replaceWithRemoteButton");
  elements.pushLocalDataButton = document.getElementById("pushLocalDataButton");
  elements.exportBackupButton = document.getElementById("exportBackupButton");
  elements.importBackupButton = document.getElementById("importBackupButton");
  elements.backupFileInput = document.getElementById("backupFileInput");
  elements.connectionStatus = document.getElementById("connectionStatus");
  elements.lastSyncText = document.getElementById("lastSyncText");
  elements.transactionModal = document.getElementById("transactionModal");
  elements.transactionModalTitle = document.getElementById("transactionModalTitle");
  elements.closeTransactionModalButton = document.getElementById("closeTransactionModalButton");
  elements.transactionForm = document.getElementById("transactionForm");
  elements.transactionDate = document.getElementById("transactionDate");
  elements.transactionCard = document.getElementById("transactionCard");
  elements.transactionAmount = document.getElementById("transactionAmount");
  elements.transactionCategory = document.getElementById("transactionCategory");
  elements.transactionDescription = document.getElementById("transactionDescription");
  elements.transactionMerchant = document.getElementById("transactionMerchant");
  elements.transactionNature = document.getElementById("transactionNature");
  elements.transactionType = document.getElementById("transactionType");
  elements.paymentStatus = document.getElementById("paymentStatus");
  elements.billPdfLink = document.getElementById("billPdfLink");
  elements.unbilledAmount = document.getElementById("unbilledAmount");
  elements.transactionNotes = document.getElementById("transactionNotes");
  elements.transactionPreview = document.getElementById("transactionPreview");
  elements.cancelTransactionEditButton = document.getElementById("cancelTransactionEditButton");
  elements.merchantSuggestions = document.getElementById("merchantSuggestions");
  elements.descriptionSuggestions = document.getElementById("descriptionSuggestions");
  elements.toast = document.getElementById("toast");
}

function bindEvents() {
  elements.viewLinks.forEach((button) => button.addEventListener("click", () => activateView(button.dataset.view)));
  elements.openTransactionModalButton.addEventListener("click", () => openTransactionModal());
  elements.mobileNewTransactionButton.addEventListener("click", () => openTransactionModal());
  elements.openMobileNavButton.addEventListener("click", openMobileNav);
  elements.closeMobileNavButton.addEventListener("click", closeMobileNav);
  elements.openCardPanelButton.addEventListener("click", () => {
    activateView("cards-view");
    elements.cardName.focus();
  });
  elements.themeToggleButton.addEventListener("click", toggleTheme);
  elements.backToDashboardButton.addEventListener("click", () => activateView("dashboard-view"));
  elements.closeTransactionModalButton.addEventListener("click", closeTransactionModal);
  elements.transactionModal.addEventListener("click", (event) => {
    if (event.target === elements.transactionModal) {
      closeTransactionModal();
    }
  });

  elements.transactionForm.addEventListener("submit", handleTransactionSubmit);
  elements.cardForm.addEventListener("submit", handleCardSubmit);
  elements.sharedLimitGroupForm.addEventListener("submit", handleSharedLimitGroupSubmit);
  elements.rewardRuleForm.addEventListener("submit", handleRewardRuleSubmit);
  elements.rewardRuleScope.addEventListener("change", updateRewardRuleScopeUI);
  elements.milestoneForm.addEventListener("submit", handleMilestoneSubmit);
  elements.settingsForm.addEventListener("submit", handleSettingsSubmit);
  elements.cancelCardEditButton.addEventListener("click", resetCardForm);
  elements.cancelSharedLimitGroupEditButton.addEventListener("click", resetSharedLimitGroupForm);
  elements.cancelTransactionEditButton.addEventListener("click", resetTransactionForm);
  elements.cancelRewardRuleEditButton.addEventListener("click", resetRewardRuleForm);
  elements.cancelMilestoneEditButton.addEventListener("click", resetMilestoneForm);
  elements.syncButton.addEventListener("click", syncFromRemote);
  elements.replaceWithRemoteButton.addEventListener("click", replaceLocalWithRemote);
  elements.pushLocalDataButton.addEventListener("click", pushLocalDataToGoogleSheets);
  elements.exportBackupButton.addEventListener("click", exportBackup);
  elements.importBackupButton.addEventListener("click", () => elements.backupFileInput.click());
  elements.backupFileInput.addEventListener("change", importBackup);

  [elements.transactionAmount, elements.transactionCategory, elements.transactionCard, elements.transactionMerchant, elements.transactionNature].forEach((field) => {
    field.addEventListener("input", updateTransactionPreview);
  });
  [elements.transactionCategory, elements.transactionCard, elements.transactionNature].forEach((field) => {
    field.addEventListener("change", updateTransactionPreview);
  });

  [
    elements.transactionSearchMerchant,
    elements.transactionSearchDescription,
    elements.transactionSearchCard,
    elements.transactionSearchCategory,
    elements.transactionSearchNature,
    elements.transactionSearchFromDate,
    elements.transactionSearchToDate,
    elements.transactionSearchMinAmount,
    elements.transactionSearchMaxAmount
  ].forEach((field) => {
    field.addEventListener("input", renderTransactionsView);
  });
  [
    elements.transactionSearchCard,
    elements.transactionSearchCategory,
    elements.transactionSearchNature
  ].forEach((field) => {
    field.addEventListener("change", renderTransactionsView);
  });
  elements.clearTransactionSearchButton.addEventListener("click", clearTransactionSearch);

  [elements.accountSearchMerchant, elements.accountSearchDescription].forEach((field) => {
    field.addEventListener("input", renderAccountView);
  });
  elements.clearAccountSearchButton.addEventListener("click", clearAccountSearch);
  [elements.cardSearchName, elements.cardSearchGroup, elements.cardSearchCycle].forEach((field) => {
    field.addEventListener("input", renderCardsTable);
  });
  elements.clearCardSearchButton.addEventListener("click", clearCardSearch);

  elements.cardTableBody.addEventListener("click", handleCardTableActions);
  elements.sharedLimitGroupsTableBody.addEventListener("click", handleSharedLimitGroupTableActions);
  elements.transactionTableBody.addEventListener("click", handleTransactionTableActions);
  elements.accountTransactionTableBody.addEventListener("click", handleTransactionTableActions);
  elements.navigatorCardList.addEventListener("click", handleOpenCardRequest);
  elements.priorityDueList.addEventListener("click", handleOpenCardRequest);
  elements.statementOrderList.addEventListener("click", handleOpenCardRequest);
  elements.utilizationList.addEventListener("click", handleOpenCardRequest);
  elements.bankGroupList.addEventListener("click", handleOpenCardRequest);
  elements.rewardRulesTableBody.addEventListener("click", handleRewardRuleTableActions);
  elements.milestonesTableBody.addEventListener("click", handleMilestoneTableActions);
  window.addEventListener("resize", handleViewportChange);
}

function loadLocalState() {
  const saved = localStorage.getItem(APP_STORAGE_KEY);
  if (!saved) {
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    state.cards = Array.isArray(parsed.cards) ? parsed.cards.map(normalizeCard) : [];
    state.transactions = Array.isArray(parsed.transactions) ? parsed.transactions.map(normalizeTransaction) : [];
    state.rewardActivities = Array.isArray(parsed.rewardActivities) ? parsed.rewardActivities.map(normalizeRewardActivity) : [];
    state.rewardRulesCatalog = Array.isArray(parsed.rewardRulesCatalog) ? parsed.rewardRulesCatalog.map(normalizeAdvancedRewardRule) : [];
    state.milestoneCatalog = Array.isArray(parsed.milestoneCatalog) ? parsed.milestoneCatalog.map(normalizeMilestoneRule) : [];
    state.sharedLimitGroups = Array.isArray(parsed.sharedLimitGroups) ? parsed.sharedLimitGroups.map(normalizeSharedLimitGroup) : [];
    state.lastSync = parsed.lastSync || null;
    dedupeState();
  } catch (error) {
    console.error("Unable to load local state", error);
  }
}

function saveLocalState() {
  localStorage.setItem(APP_STORAGE_KEY, JSON.stringify({
    cards: state.cards,
    transactions: state.transactions,
    rewardActivities: state.rewardActivities,
    rewardRulesCatalog: state.rewardRulesCatalog,
    milestoneCatalog: state.milestoneCatalog,
    sharedLimitGroups: state.sharedLimitGroups,
    lastSync: state.lastSync
  }));
}

function seedStarterData() {
  if (state.cards.length > 0) {
    return;
  }

  state.cards = STARTER_CARD_NAMES.map((cardName, index) => {
    const bankName = inferBankName(cardName);
    const statementDate = ((index * 3) % 28) + 1;
    const dueDate = (((index * 3) + 18) % 28) + 1;
    return normalizeCard({
      id: createId(),
      cardName,
      bankName,
      creditLimit: 100000 + (index * 17000),
      billingCycleDate: statementDate,
      dueDate,
      statementDate,
      anniversaryDate: new Date().toISOString().split("T")[0],
      renewalTarget: 300000,
      sharedGroup: bankName,
      preferredCategories: getDefaultCategoriesForCard(cardName),
      preferredTransactionTypes: ["Online", "Offline / POS", "UPI"],
      supportedTransactionNature: ["Expense", "Payment", "Refund"],
      rewardRules: getDefaultRewardRulesForCard(cardName),
      advancedRewardRules: getDefaultAdvancedRulesForCard(cardName),
      milestones: [],
      annualBenefits: []
    });
  });

  saveLocalState();
}

function dedupeState() {
  state.cards = dedupeCardsByName(state.cards, state.transactions);
  state.rewardRulesCatalog = dedupeItemsByKey(state.rewardRulesCatalog, (rule) => [
    normalizeCardKey(rule.cardName),
    normalizeCardKey(rule.category),
    normalizeCardKey(rule.scope),
    normalizeCardKey(rule.merchant),
    Number(rule.baseUnit || 0),
    Number(rule.rewardPerUnit || 0),
    Number(rule.multiplier || 0),
    Number(rule.priority || 0)
  ].join("|"));
  state.milestoneCatalog = dedupeItemsByKey(state.milestoneCatalog, (rule) => [
    normalizeCardKey(rule.cardName),
    normalizeCardKey(rule.cycleType),
    normalizeCardKey(rule.milestoneType),
    Number(rule.target || 0),
    normalizeCardKey(rule.rewardType),
    Number(rule.rewardValue || 0)
  ].join("|"));
  state.sharedLimitGroups = dedupeItemsByKey(state.sharedLimitGroups, (group) => normalizeCardKey(group.groupName));

  if (state.ui.activeCardId && !getCardById(state.ui.activeCardId)) {
    state.ui.activeCardId = null;
  }
}

function removeStarterCardsFromState(localCards, remoteCards) {
  if (!Array.isArray(remoteCards) || !remoteCards.length) {
    return localCards || [];
  }

  return (localCards || []).filter((card) => {
    const cardNameKey = normalizeCardKey(card.cardName);
    const isStarterCard = STARTER_CARD_NAMES.some((name) => normalizeCardKey(name) === cardNameKey);
    if (!isStarterCard) {
      return true;
    }

    const existsInRemote = remoteCards.some((remoteCard) => normalizeCardKey(remoteCard.cardName) === cardNameKey);
    return !existsInRemote;
  });
}

function dedupeCardsByName(cards, transactions) {
  const deduped = new Map();

  (cards || []).forEach((card) => {
    const normalized = normalizeCard(card);
    const key = normalizeCardKey(normalized.cardName);
    if (!key) {
      return;
    }

    const existing = deduped.get(key);
    if (!existing) {
      deduped.set(key, normalized);
      return;
    }

    const existingScore = getCardMergeScore(existing, transactions);
    const currentScore = getCardMergeScore(normalized, transactions);
    deduped.set(key, currentScore >= existingScore ? normalized : existing);
  });

  return Array.from(deduped.values());
}

function getCardMergeScore(card, transactions) {
  const linkedTransactions = (transactions || []).filter((item) => {
    return String(item.cardId || "") === String(card.id)
      || normalizeCardKey(item.cardName) === normalizeCardKey(card.cardName);
  }).length;

  const richness = [
    Number(card.creditLimit || 0) > 0,
    Number(card.statementDate || 0) > 0,
    Number(card.dueDate || 0) > 0,
    Boolean(card.bankName),
    Boolean(card.anniversaryDate),
    Array.isArray(card.preferredCategories) && card.preferredCategories.length > 0,
    Array.isArray(card.preferredTransactionTypes) && card.preferredTransactionTypes.length > 0
  ].filter(Boolean).length;

  return (linkedTransactions * 100) + richness;
}

function dedupeItemsByKey(items, getKey) {
  const map = new Map();
  (items || []).forEach((item) => {
    const normalized = item;
    const key = getKey(normalized);
    if (!key) {
      return;
    }
    map.set(key, normalized);
  });
  return Array.from(map.values());
}

function getRewardRulesForCard(card) {
  const cardNameKey = normalizeCardKey(card.cardName);
  const sheetRules = state.rewardRulesCatalog
    .filter((rule) => normalizeCardKey(rule.cardName) === cardNameKey)
    .map(normalizeAdvancedRewardRule);

  if (sheetRules.length) {
    return sheetRules;
  }

  return Array.isArray(card.advancedRewardRules)
    ? card.advancedRewardRules.map(normalizeAdvancedRewardRule)
    : [];
}

function getMilestonesForCard(card) {
  const cardNameKey = normalizeCardKey(card.cardName);
  const catalogRules = state.milestoneCatalog
    .filter((rule) => normalizeCardKey(rule.cardName) === cardNameKey)
    .map(normalizeMilestoneRule);

  if (catalogRules.length) {
    return catalogRules;
  }

  const fallback = Array.isArray(card.annualBenefits) && card.annualBenefits.length ? card.annualBenefits : card.milestones;
  return Array.isArray(fallback) ? fallback.map(normalizeMilestoneRule) : [];
}

function getCycleSpendForMilestone(card, cycleType) {
  const normalizedCycle = String(cycleType || "Annual").trim().toLowerCase();
  const transactions = getCardExpenseTransactions(card);
  const today = new Date();

  if (normalizedCycle === "monthly") {
    return transactions
      .filter((item) => {
        const date = new Date(item.date);
        return date.getFullYear() === today.getFullYear() && date.getMonth() === today.getMonth();
      })
      .reduce((sum, item) => sum + Number(item.amount || 0), 0);
  }

  if (normalizedCycle === "quarterly") {
    const currentQuarter = Math.floor(today.getMonth() / 3);
    return transactions
      .filter((item) => {
        const date = new Date(item.date);
        return date.getFullYear() === today.getFullYear() && Math.floor(date.getMonth() / 3) === currentQuarter;
      })
      .reduce((sum, item) => sum + Number(item.amount || 0), 0);
  }

  return getCardAnniversarySpend(card);
}

function setDefaultDates() {
  const today = new Date().toISOString().split("T")[0];
  elements.transactionDate.value = today;
  if (!elements.anniversaryDate.value) {
    elements.anniversaryDate.value = today;
  }
  elements.scriptUrl.value = state.settings.scriptUrl;
}

function populateStaticSelects() {
  elements.transactionCategory.innerHTML = buildOptions(CATEGORIES);
  elements.transactionSearchCategory.innerHTML = `<option value="">All Categories</option>${buildOptions(CATEGORIES)}`;
  elements.transactionNature.innerHTML = buildOptions(TRANSACTION_NATURES);
  elements.transactionSearchNature.innerHTML = `<option value="">All Types</option>${buildOptions(TRANSACTION_NATURES)}`;
  elements.transactionType.innerHTML = buildOptions(TRANSACTION_TYPES);
  elements.paymentStatus.innerHTML = buildOptions(PAYMENT_STATUSES);
  populateRewardRuleCategorySuggestions();
}

function populateRewardRuleCategorySuggestions() {
  const savedRuleCategories = state.rewardRulesCatalog
    .map((item) => String(item.category || "").trim())
    .filter(Boolean);
  const uniqueCategories = [...new Set([...REWARD_RULE_CATEGORIES, ...savedRuleCategories])].sort((a, b) => a.localeCompare(b));
  elements.rewardRuleCategorySuggestions.innerHTML = uniqueCategories
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");

  const savedMerchants = state.rewardRulesCatalog
    .map((item) => String(item.merchant || "").trim())
    .filter(Boolean);
  const uniqueMerchants = [...new Set(["Any", ...savedMerchants])].sort((a, b) => a.localeCompare(b));
  elements.rewardRuleMerchantSuggestions.innerHTML = uniqueMerchants
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function populateCardDropdowns() {
  const sortedCards = getSortedCards();
  const transactionOptions = sortedCards.length
    ? sortedCards.map((card) => `<option value="${card.id}">${escapeHtml(card.cardName)}</option>`).join("")
    : `<option value="">No cards available</option>`;
  const searchOptions = sortedCards.length
    ? `<option value="">All Cards</option>${sortedCards.map((card) => `<option value="${card.id}">${escapeHtml(card.cardName)}</option>`).join("")}`
    : `<option value="">All Cards</option>`;

  elements.transactionCard.innerHTML = transactionOptions;
  elements.transactionSearchCard.innerHTML = searchOptions;
  elements.rewardRuleCard.innerHTML = transactionOptions;
  elements.milestoneCard.innerHTML = transactionOptions;

  if (sortedCards.length) {
    const activeCard = getCardById(state.ui.activeCardId);
    const preferredCardId = state.ui.editingTransactionId ? elements.transactionCard.value : (activeCard ? activeCard.id : sortedCards[0].id);
    elements.transactionCard.value = getCardById(preferredCardId) ? preferredCardId : sortedCards[0].id;
    if (!getCardById(elements.transactionSearchCard.value)) {
      elements.transactionSearchCard.value = "";
    }
    if (!getCardById(elements.rewardRuleCard.value)) {
      elements.rewardRuleCard.value = sortedCards[0].id;
    }
    if (!getCardById(elements.milestoneCard.value)) {
      elements.milestoneCard.value = sortedCards[0].id;
    }
  }

  updateTransactionPreview();
}

function renderApp() {
  populateCardDropdowns();
  populateSharedLimitGroupMemberOptions();
  renderConnectionStatus();
  renderThemeToggle();
  populateRewardRuleCategorySuggestions();
  renderNavigator();
  renderDashboard();
  renderRewardRulesManager();
  renderMilestonesManager();
  renderTransactionsView();
  renderCardsTable();
  renderSharedLimitGroupsTable();
  renderAccountView();
  renderTransactionSuggestions();
  updateTransactionPreview();
  renderViews();
}

function applySavedTheme() {
  const savedTheme = localStorage.getItem(THEME_KEY) || "light";
  document.body.classList.toggle("dark-theme", savedTheme === "dark");
}

function toggleTheme() {
  const isDark = document.body.classList.toggle("dark-theme");
  localStorage.setItem(THEME_KEY, isDark ? "dark" : "light");
  renderThemeToggle();
}

function renderThemeToggle() {
  if (!elements.themeToggleButton) {
    return;
  }
  elements.themeToggleButton.textContent = document.body.classList.contains("dark-theme") ? "Light" : "Dark";
}

function isMobileViewport() {
  return window.matchMedia("(max-width: 900px)").matches;
}

function openMobileNav() {
  state.ui.mobileNavOpen = true;
  renderMobileNavState();
}

function closeMobileNav() {
  state.ui.mobileNavOpen = false;
  renderMobileNavState();
}

function handleViewportChange() {
  if (!isMobileViewport()) {
    state.ui.mobileNavOpen = false;
  }
  renderMobileNavState();
}

function renderMobileNavState() {
  if (!elements.navigatorPanel) {
    return;
  }
  elements.navigatorPanel.classList.toggle("mobile-open", state.ui.mobileNavOpen && isMobileViewport());
  document.body.classList.toggle("mobile-nav-open", state.ui.mobileNavOpen && isMobileViewport());
}

function activateView(viewId) {
  state.ui.activeView = viewId;
  if (isMobileViewport()) {
    closeMobileNav();
  }
  renderViews();
  if (viewId === "transactions-view") {
    renderTransactionsView();
  }
  if (viewId === "account-view") {
    renderAccountView();
  }
}

function renderViews() {
  elements.viewLinks.forEach((button) => button.classList.toggle("active", button.dataset.view === state.ui.activeView));
  elements.workspaceViews.forEach((section) => section.classList.toggle("active", section.id === state.ui.activeView));
  if (elements.mobileActiveViewLabel) {
    const activeLink = elements.viewLinks.find((button) => button.dataset.view === state.ui.activeView);
    elements.mobileActiveViewLabel.textContent = activeLink ? activeLink.textContent.trim() : "Account";
  }
}

function renderNavigator() {
  const sortedCards = getSortedCards();
  elements.navigatorCardCount.textContent = String(sortedCards.length);

  if (!sortedCards.length) {
    elements.navigatorCardList.className = "navigator-cards empty-state";
    elements.navigatorCardList.textContent = "No cards added yet.";
    return;
  }

  elements.navigatorCardList.className = "navigator-cards";
  elements.navigatorCardList.innerHTML = sortedCards.map((card) => `
    <button class="navigator-card ${state.ui.activeCardId === card.id ? "active" : ""}" type="button" data-open-card="${card.id}">
      <strong>${escapeHtml(card.cardName)}</strong>
    </button>
  `).join("");
}

function renderDashboard() {
  const cards = getSortedCards();
  const alertData = getAlertData();
  const dueSoon = alertData.filter((item) => item.daysLeft >= 0 && item.daysLeft <= 5 && item.outstanding > 0);

  elements.dashboardCardCount.textContent = String(cards.length);
  elements.upcomingDues.textContent = String(dueSoon.length);
  renderHighSpendAlert(cards);
  renderPriorityDueList(dueSoon);
  renderStatementOrder(cards);
  renderUtilization(cards);
  renderBankGroups(cards);
}

function renderHighSpendAlert(cards) {
  const highSpendCards = cards
    .map((card) => ({
      card,
      spend: getCardAnniversarySpend(card)
    }))
    .filter((item) => item.spend > 800000)
    .sort((a, b) => b.spend - a.spend);

  if (!highSpendCards.length) {
    elements.highSpendAlertStrip.className = "panel empty-state";
    elements.highSpendAlertStrip.textContent = "No high annual-spend cards right now.";
    return;
  }

  elements.highSpendAlertStrip.className = "panel high-spend-strip";
  elements.highSpendAlertStrip.innerHTML = `
    <div class="row-between">
      <div>
        <h3>High Anniversary Spend Alert</h3>
        <p class="mini-text">Cards with anniversary-year spend above ₹8,00,000.</p>
      </div>
      <span class="pill danger">${highSpendCards.length} card${highSpendCards.length === 1 ? "" : "s"}</span>
    </div>
    <div class="high-spend-list">
      ${highSpendCards.map((item) => `
        <button class="high-spend-chip" type="button" data-open-card="${item.card.id}">
          <span>${escapeHtml(item.card.cardName)}</span>
          <span>${formatCurrency(item.spend)}</span>
        </button>
      `).join("")}
    </div>
  `;
}

function renderRewardRulesManager() {
  const rules = [...state.rewardRulesCatalog].sort((a, b) => {
    const cardDiff = String(a.cardName || "").localeCompare(String(b.cardName || ""));
    if (cardDiff !== 0) {
      return cardDiff;
    }
    return Number(a.priority || 999) - Number(b.priority || 999);
  });

  if (!rules.length) {
    elements.rewardRulesTableBody.innerHTML = `<tr><td colspan="10">No reward rules added yet.</td></tr>`;
    return;
  }

  elements.rewardRulesTableBody.innerHTML = rules.map((rule) => `
    <tr>
      <td>${escapeHtml(rule.cardName || "-")}</td>
      <td>${escapeHtml(rule.category || "-")}</td>
      <td>${escapeHtml(rule.scope || "category-merchant")}</td>
      <td>${escapeHtml(rule.merchant || "-")}</td>
      <td>${rule.baseUnit || "-"}</td>
      <td>${rule.rewardPerUnit || 0}</td>
      <td>${rule.multiplier || 1}</td>
      <td>${Number(rule.rewardPerUnit || 0) * Number(rule.multiplier || 1)}</td>
      <td>${rule.priority || 999}</td>
      <td><button class="table-action-btn" type="button" data-edit-reward-rule="${rule.id}">Edit</button></td>
    </tr>
  `).join("");
}

function renderMilestonesManager() {
  const rules = [...state.milestoneCatalog].sort((a, b) => {
    const cardDiff = String(a.cardName || "").localeCompare(String(b.cardName || ""));
    if (cardDiff !== 0) {
      return cardDiff;
    }
    return Number(a.target || 0) - Number(b.target || 0);
  });

  if (!rules.length) {
    elements.milestonesTableBody.innerHTML = `<tr><td colspan="8">No milestones added yet.</td></tr>`;
    return;
  }

  elements.milestonesTableBody.innerHTML = rules.map((rule) => `
    <tr>
      <td>${escapeHtml(rule.cardName || "-")}</td>
      <td>${escapeHtml(rule.cycleType || "Annual")}</td>
      <td>${escapeHtml(rule.milestoneType || "-")}</td>
      <td>${formatCurrency(rule.target || 0)}</td>
      <td>${escapeHtml(rule.rewardType || "-")}</td>
      <td>${formatCurrency(rule.rewardValue || 0)}</td>
      <td>${escapeHtml(rule.notes || "-")}</td>
      <td><button class="table-action-btn" type="button" data-edit-milestone="${rule.id}">Edit</button></td>
    </tr>
  `).join("");
}

function renderPriorityDueList(items) {
  if (!items.length) {
    elements.priorityDueList.className = "stack-list empty-state";
    elements.priorityDueList.textContent = "No card is due within the next 5 days.";
    return;
  }

  elements.priorityDueList.className = "stack-list";
  elements.priorityDueList.innerHTML = items.map((item) => `
    <article class="list-card clickable-card" data-open-card="${item.cardId}">
      <div class="row-between">
        <h4>${escapeHtml(item.cardName)}</h4>
        <span class="pill ${item.daysLeft <= 1 ? "warn" : "good"}">${item.daysLeft} day${item.daysLeft === 1 ? "" : "s"}</span>
      </div>
      <p class="mini-text">Due ${formatDate(item.liveDueDate)} | Outstanding ${formatCurrency(item.outstanding)}</p>
    </article>
  `).join("");
}

function renderStatementOrder(cards) {
  if (!cards.length) {
    elements.statementOrderList.className = "stack-list empty-state";
    elements.statementOrderList.textContent = "No cards available yet.";
    return;
  }

  const sorted = [...cards].sort((a, b) => (Number(a.statementDate || 99) - Number(b.statementDate || 99)) || a.cardName.localeCompare(b.cardName));
  elements.statementOrderList.className = "stack-list";
  elements.statementOrderList.innerHTML = sorted.map((card, index) => `
    <article class="list-card clickable-card" data-open-card="${card.id}">
      <div class="row-between">
        <h4>${index + 1}. ${escapeHtml(card.cardName)}</h4>
        <strong>${formatOrdinalDay(card.statementDate)}</strong>
      </div>
      <p class="mini-text">${escapeHtml(card.bankName)} | Due ${card.dueDate || "-"} | Group ${escapeHtml(card.sharedGroup || card.bankName || "-")}</p>
    </article>
  `).join("");
}

function renderUtilization(cards) {
  if (!cards.length) {
    elements.utilizationList.className = "stack-list empty-state";
    elements.utilizationList.textContent = "No cards available yet.";
    return;
  }

  elements.utilizationList.className = "stack-list";
  elements.utilizationList.innerHTML = cards.map((card) => {
    const sharedGroup = getSharedLimitGroupForCard(card);
    const utilized = sharedGroup ? getSharedGroupNetUtilizedAmount(sharedGroup.groupName) : getCardNetUtilizedAmount(card);
    const limit = getEffectiveCardLimit(card);
    const utilization = limit > 0 ? Math.min((utilized / limit) * 100, 999) : 0;
    return `
      <article class="list-card clickable-card" data-open-card="${card.id}">
        <div class="row-between">
          <h4>${escapeHtml(card.cardName)}</h4>
          <strong>${utilization.toFixed(1)}%</strong>
        </div>
        <p class="mini-text">Used ${formatCurrency(utilized)} of ${formatCurrency(limit)} | Available ${formatCurrency(getCardAvailableLimit(card))}${sharedGroup ? ` | Shared ${escapeHtml(sharedGroup.groupName)}` : ""}</p>
      </article>
    `;
  }).join("");
}

function renderBankGroups(cards) {
  if (!cards.length) {
    elements.bankGroupList.className = "bank-group-list empty-state";
    elements.bankGroupList.textContent = "No grouped card data yet.";
    return;
  }

  const grouped = {};
  cards.forEach((card) => {
    const key = card.bankName || "Others";
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(card);
  });

  const bankNames = Object.keys(grouped).sort((a, b) => a.localeCompare(b));
  elements.bankGroupList.className = "bank-group-list";
  elements.bankGroupList.innerHTML = bankNames.map((bankName) => `
    <article class="bank-card">
      <div class="row-between">
        <h4>${escapeHtml(bankName)}</h4>
        <span class="pill good">${grouped[bankName].length} card${grouped[bankName].length === 1 ? "" : "s"}</span>
      </div>
      <div class="bank-card-list">
        ${grouped[bankName].sort((a, b) => a.cardName.localeCompare(b.cardName)).map((card) => `
          <div class="bank-row clickable-card" data-open-card="${card.id}">
            <div>
              <strong>${escapeHtml(card.cardName)}</strong>
              <div class="mini-text">Due ${card.dueDate || "-"} | Stmt ${card.statementDate || "-"}</div>
            </div>
            <div class="mini-text">${formatCurrency(getCardAvailableLimit(card))}</div>
          </div>
        `).join("")}
      </div>
    </article>
  `).join("");
}
function renderTransactionsView() {
  const filtered = getAllTransactionFilters().reduce((items, filterFn) => items.filter(filterFn), getSortedTransactions());
  renderTransactionSearchSummary(filtered);

  if (!filtered.length) {
    elements.transactionTableBody.innerHTML = `<tr><td colspan="8">No transactions recorded yet.</td></tr>`;
    return;
  }

  elements.transactionTableBody.innerHTML = filtered.map((transaction) => `
    <tr>
      <td>${formatDate(transaction.date)}</td>
      <td>${escapeHtml(transaction.cardName || getCardNameById(transaction.cardId))}</td>
      <td>
        <strong>${escapeHtml(transaction.merchant || "-")}</strong>
        <div class="mini-text">${escapeHtml(transaction.description || "-")}</div>
      </td>
      <td>${escapeHtml(transaction.category || "-")}</td>
      <td>
        <div class="semantic-badge ${getTransactionTone(transaction.transactionNature)}">${escapeHtml(getNatureLabel(transaction.transactionNature))}</div>
        <div class="mini-text">${escapeHtml(transaction.transactionType || "-")}</div>
      </td>
      <td><span class="${getAmountClass(transaction)}">${formatCurrency(transaction.amount)}</span></td>
      <td>${formatPoints(transaction.rewardPoints)}</td>
      <td><button class="table-action-btn" type="button" data-edit-transaction="${transaction.id}">Edit</button></td>
    </tr>
  `).join("");
}

function renderTransactionSearchSummary(filtered) {
  const merchantQuery = normalizeCardKey(elements.transactionSearchMerchant.value);
  const descriptionQuery = normalizeCardKey(elements.transactionSearchDescription.value);
  const cardLabel = getCardNameById(elements.transactionSearchCard.value);
  const category = elements.transactionSearchCategory.value.trim();
  const nature = elements.transactionSearchNature.value.trim();
  const fromDate = elements.transactionSearchFromDate.value;
  const toDate = elements.transactionSearchToDate.value;
  const minAmount = elements.transactionSearchMinAmount.value.trim();
  const maxAmount = elements.transactionSearchMaxAmount.value.trim();

  if (!merchantQuery && !descriptionQuery && !elements.transactionSearchCard.value && !category && !nature && !fromDate && !toDate && !minAmount && !maxAmount) {
    elements.transactionSearchSummary.className = "result-card empty-state";
    elements.transactionSearchSummary.textContent = "Search by client, merchant, card, date range, amount, or transaction type to see matching totals across all cards.";
    return;
  }

  const total = filtered.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const rewardTotal = filtered.reduce((sum, item) => sum + Number(item.rewardPoints || 0), 0);
  const cardsUsed = [...new Set(filtered.map((item) => item.cardName).filter(Boolean))];
  const merchants = [...new Set(filtered.map((item) => item.merchant).filter(Boolean))];
  const clients = [...new Set(filtered.map((item) => item.description).filter(Boolean))];

  elements.transactionSearchSummary.className = "result-card";
  elements.transactionSearchSummary.innerHTML = `
    <div class="row-between">
      <div>
        <h4>${filtered.length} matching transaction${filtered.length === 1 ? "" : "s"}</h4>
        <p class="mini-text">${[
          merchantQuery ? `Merchant: ${escapeHtml(elements.transactionSearchMerchant.value.trim())}` : "",
          descriptionQuery ? `Client: ${escapeHtml(elements.transactionSearchDescription.value.trim())}` : "",
          cardLabel ? `Card: ${escapeHtml(cardLabel)}` : "",
          category ? `Category: ${escapeHtml(category)}` : "",
          nature ? `Type: ${escapeHtml(nature)}` : "",
          fromDate ? `From: ${formatDate(fromDate)}` : "",
          toDate ? `To: ${formatDate(toDate)}` : "",
          minAmount ? `Min: ${formatCurrency(minAmount)}` : "",
          maxAmount ? `Max: ${formatCurrency(maxAmount)}` : ""
        ].filter(Boolean).join(" | ")}</p>
      </div>
      <span class="pill good">${cardsUsed.length} card${cardsUsed.length === 1 ? "" : "s"}</span>
    </div>
    <p><strong>Total amount:</strong> ${formatCurrency(total)}</p>
    <p><strong>Total rewards:</strong> ${formatPoints(rewardTotal)}</p>
    <p><strong>Cards used:</strong> ${escapeHtml(cardsUsed.join(", ") || "-")}</p>
    <p><strong>Clients found:</strong> ${escapeHtml(clients.join(", ") || "-")}</p>
    <p><strong>Merchants found:</strong> ${escapeHtml(merchants.join(", ") || "-")}</p>
  `;
}

function clearTransactionSearch() {
  elements.transactionSearchMerchant.value = "";
  elements.transactionSearchDescription.value = "";
  elements.transactionSearchCard.value = "";
  elements.transactionSearchCategory.value = "";
  elements.transactionSearchNature.value = "";
  elements.transactionSearchFromDate.value = "";
  elements.transactionSearchToDate.value = "";
  elements.transactionSearchMinAmount.value = "";
  elements.transactionSearchMaxAmount.value = "";
  renderTransactionsView();
}

function renderAccountView() {
  const card = getCardById(state.ui.activeCardId);

  if (!card) {
    elements.accountViewTitle.textContent = "Account View";
    elements.accountViewSubtitle.textContent = "Select a card account from the navigator to open its ledger.";
    elements.accountAvailableLimit.textContent = "₹0";
    elements.accountTotalLimit.textContent = "₹0";
    elements.accountOutstanding.textContent = "₹0";
    elements.accountUnbilled.textContent = "₹0";
    elements.accountAnniversarySpend.textContent = "₹0";
    elements.accountRewardTotal.textContent = "0 pts";
    elements.accountCycleLabel.textContent = "-";
    elements.accountCycleBilled.textContent = "₹0";
    elements.accountCyclePaid.textContent = "₹0";
    elements.accountCyclePending.textContent = "₹0";
    elements.accountMetaSummary.className = "stack-list empty-state";
    elements.accountMetaSummary.textContent = "No card selected.";
    elements.accountSearchSummary.className = "result-card empty-state";
    elements.accountSearchSummary.textContent = "Search this card ledger by merchant or client.";
    elements.accountTransactionTableBody.innerHTML = `<tr><td colspan="8">No card selected.</td></tr>`;
    return;
  }

  const accountTransactions = getCardTransactions(card);
  const filteredTransactions = getFilteredAccountTransactions(card);

  elements.accountViewTitle.textContent = card.cardName;
  elements.accountViewSubtitle.textContent = `${card.bankName} | Stmt ${card.statementDate || "-"} | Due ${card.dueDate || "-"} | Group ${card.sharedGroup || card.bankName || "-"}`;
  elements.accountAvailableLimit.textContent = formatCurrency(getCardAvailableLimit(card));
  elements.accountTotalLimit.textContent = formatCurrency(getEffectiveCardLimit(card));
  elements.accountOutstanding.textContent = formatCurrency(getCardOutstanding(card));
  elements.accountUnbilled.textContent = formatCurrency(getCardUnbilledAmount(card));
  elements.accountAnniversarySpend.textContent = formatCurrency(getCardAnniversarySpend(card));
  elements.accountRewardTotal.textContent = formatPoints(getCardRewardTotal(card));
  renderAccountCycleSummary(card);

  renderAccountMetaSummary(card, accountTransactions);
  renderAccountSearchSummary(card, filteredTransactions);

  if (!filteredTransactions.length) {
    elements.accountTransactionTableBody.innerHTML = `<tr><td colspan="8">No transactions match this account search.</td></tr>`;
    return;
  }

  elements.accountTransactionTableBody.innerHTML = filteredTransactions.map((transaction) => `
    <tr>
      <td>${formatDate(transaction.date)}</td>
      <td>
        <strong>${escapeHtml(transaction.merchant || "-")}</strong>
        <div class="mini-text">${escapeHtml(transaction.description || "-")}</div>
      </td>
      <td>${escapeHtml(transaction.category || "-")}</td>
      <td><div class="semantic-badge ${getTransactionTone(transaction.transactionNature)}">${escapeHtml(getNatureLabel(transaction.transactionNature))}</div></td>
      <td><span class="${getAmountClass(transaction)}">${formatCurrency(transaction.amount)}</span></td>
      <td>${formatPoints(transaction.rewardPoints)}</td>
      <td>${escapeHtml(transaction.paymentStatus || "-")}</td>
      <td><button class="table-action-btn" type="button" data-edit-transaction="${transaction.id}">Edit</button></td>
    </tr>
  `).join("");
}

function renderAccountMetaSummary(card, transactions) {
  const nextDue = getLiveDueDate(card);
  const cycle = getLatestBillCycle(card);
  const sharedGroup = getSharedLimitGroupForCard(card);
  const sharedMembers = sharedGroup ? getCardsInSharedGroup(sharedGroup.groupName).map((item) => item.cardName).filter((name) => name !== card.cardName) : [];
  const totalTransactions = transactions.length;
  const paymentCount = transactions.filter((item) => isCreditNature(item.transactionNature)).length;
  const uniqueMerchants = [...new Set(transactions.map((item) => normalizeCardKey(item.merchant)).filter(Boolean))].length;
  const uniqueClients = [...new Set(transactions.map((item) => normalizeCardKey(item.description)).filter(Boolean))].length;
  const latest = transactions[0];

  elements.accountMetaSummary.className = "account-meta-grid";
  elements.accountMetaSummary.innerHTML = `
    <div class="meta-block">
      <strong>Billing & Limit</strong>
      <div class="summary-copy">Available ${formatCurrency(getCardAvailableLimit(card))} | Total ${formatCurrency(getEffectiveCardLimit(card))}</div>
      <div class="summary-copy">Outstanding ${formatCurrency(getCardOutstanding(card))} | Unbilled ${formatCurrency(getCardUnbilledAmount(card))}</div>
      <div class="summary-copy">Live due date ${formatDate(nextDue)}</div>
      <div class="summary-copy">${cycle.label} bill ${formatCurrency(cycle.billedAmount)} | Paid ${formatCurrency(cycle.paidAmount)} | Pending ${formatCurrency(cycle.pendingAmount)}</div>
      <div class="summary-copy">${sharedGroup ? `Shared group ${escapeHtml(sharedGroup.groupName)} | ${sharedMembers.length ? `With ${escapeHtml(sharedMembers.join(", "))}` : "Single mapped card"}` : "Standalone card limit"}</div>
    </div>
    <div class="meta-block">
      <strong>Usage Summary</strong>
      <div class="summary-copy">${totalTransactions} total transaction${totalTransactions === 1 ? "" : "s"} | ${paymentCount} credit/payment entries</div>
      <div class="summary-copy">${uniqueMerchants} merchants | ${uniqueClients} client descriptions</div>
      <div class="summary-copy">Anniversary spend ${formatCurrency(getCardAnniversarySpend(card))}</div>
    </div>
    <div class="meta-block">
      <strong>Rewards & Milestones</strong>
      <div class="summary-copy">Earned rewards ${formatPoints(getCardRewardEarnedFromTransactions(card))}</div>
      <div class="summary-copy">Redeemed ${formatPoints(getCardRedeemedPoints(card))} | Expired ${formatPoints(getCardExpiredPoints(card))}</div>
      <div class="summary-copy">${getMilestoneStatus(card)}</div>
    </div>
    <div class="meta-block">
      <strong>Latest Activity</strong>
      <div class="summary-copy">${latest ? `${formatDate(latest.date)} | ${escapeHtml(latest.merchant || "-")}` : "No transactions yet."}</div>
      <div class="summary-copy">${latest ? `${escapeHtml(getNatureLabel(latest.transactionNature))} | ${formatCurrency(latest.amount)}` : ""}</div>
      <div class="summary-copy">${latest ? escapeHtml(latest.description || "-") : ""}</div>
    </div>
  `;
}

function renderAccountCycleSummary(card) {
  const cycle = getLatestBillCycle(card);
  elements.accountCycleLabel.textContent = cycle.label;
  elements.accountCycleBilled.textContent = formatCurrency(cycle.billedAmount);
  elements.accountCyclePaid.textContent = formatCurrency(cycle.paidAmount);
  elements.accountCyclePending.textContent = formatCurrency(cycle.pendingAmount);
}

function renderAccountSearchSummary(card, filteredTransactions) {
  const merchantQuery = normalizeCardKey(elements.accountSearchMerchant.value);
  const descriptionQuery = normalizeCardKey(elements.accountSearchDescription.value);

  if (!merchantQuery && !descriptionQuery) {
    elements.accountSearchSummary.className = "result-card empty-state";
    elements.accountSearchSummary.textContent = "Search this card ledger by merchant or client.";
    return;
  }

  const total = filteredTransactions.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const rewardTotal = filteredTransactions.reduce((sum, item) => sum + Number(item.rewardPoints || 0), 0);
  const merchants = [...new Set(filteredTransactions.map((item) => item.merchant).filter(Boolean))];
  const descriptions = [...new Set(filteredTransactions.map((item) => item.description).filter(Boolean))];

  elements.accountSearchSummary.className = "result-card";
  elements.accountSearchSummary.innerHTML = `
    <div class="row-between">
      <div>
        <h4>${escapeHtml(card.cardName)} search summary</h4>
        <p class="mini-text">${[
          merchantQuery ? `Merchant: ${escapeHtml(elements.accountSearchMerchant.value.trim())}` : "",
          descriptionQuery ? `Client: ${escapeHtml(elements.accountSearchDescription.value.trim())}` : ""
        ].filter(Boolean).join(" | ")}</p>
      </div>
      <span class="pill good">${filteredTransactions.length} match${filteredTransactions.length === 1 ? "" : "es"}</span>
    </div>
    <p><strong>Total amount:</strong> ${formatCurrency(total)}</p>
    <p><strong>Total rewards:</strong> ${formatPoints(rewardTotal)}</p>
    <p><strong>Merchants:</strong> ${escapeHtml(merchants.join(", ") || "-")}</p>
    <p><strong>Clients:</strong> ${escapeHtml(descriptions.join(", ") || "-")}</p>
  `;
}

function clearAccountSearch() {
  elements.accountSearchMerchant.value = "";
  elements.accountSearchDescription.value = "";
  renderAccountView();
}

function renderCardsTable() {
  const cards = getSortedCards();
  renderCardPortfolioSummary(cards);
  if (!cards.length) {
    elements.cardTableBody.innerHTML = `<tr><td colspan="11">No cards added yet.</td></tr>`;
    elements.cardsSearchSummary.className = "result-card empty-state";
    elements.cardsSearchSummary.textContent = "Search cards by name, bank, group, or billing cycle to review your portfolio faster.";
    return;
  }

  const filteredCards = getFilteredCards(cards);
  renderCardsSearchSummary(filteredCards, cards.length);

  if (!filteredCards.length) {
    elements.cardTableBody.innerHTML = `<tr><td colspan="11">No cards match this search.</td></tr>`;
    return;
  }

  elements.cardTableBody.innerHTML = filteredCards.map((card) => {
    const rewardRuleCount = getRewardRulesForCard(card).length;
    const milestoneCount = getMilestonesForCard(card).length;
    return `
    <tr>
      <td>${escapeHtml(card.cardName)}</td>
      <td>${escapeHtml(card.bankName)}</td>
      <td>${escapeHtml(card.sharedGroup || "-")}</td>
      <td>${formatCurrency(getEffectiveCardLimit(card))}</td>
      <td>${formatCurrency(getCardAvailableLimit(card))}</td>
      <td>${formatCurrency(getCardOutstanding(card))}</td>
      <td>${formatOrdinalDay(card.statementDate)}</td>
      <td>${formatOrdinalDay(card.dueDate)}</td>
      <td><span class="pill ${rewardRuleCount ? "good" : "warn"}">${rewardRuleCount ? `${rewardRuleCount} rule${rewardRuleCount === 1 ? "" : "s"}` : "Pending"}</span></td>
      <td><span class="pill ${milestoneCount ? "good" : "warn"}">${milestoneCount ? `${milestoneCount} item${milestoneCount === 1 ? "" : "s"}` : "Pending"}</span></td>
      <td>
        <div class="table-action-row">
          <button class="table-action-btn" type="button" data-open-card-row="${card.id}">Open</button>
          <button class="table-action-btn" type="button" data-edit-card="${card.id}">Edit</button>
        </div>
      </td>
    </tr>
  `;
  }).join("");
}

function renderCardPortfolioSummary(cards) {
  const bankCount = new Set(cards.map((card) => normalizeCardKey(card.bankName)).filter(Boolean)).size;
  const sharedCount = state.sharedLimitGroups.length;
  const dueSoonCount = getAlertData().filter((item) => item.daysLeft >= 0 && item.daysLeft <= 7 && item.outstanding > 0).length;

  elements.cardSummaryTotal.textContent = String(cards.length);
  elements.cardSummaryBanks.textContent = String(bankCount);
  elements.cardSummaryShared.textContent = String(sharedCount);
  elements.cardSummaryDueSoon.textContent = String(dueSoonCount);
}

function getFilteredCards(cards) {
  const nameQuery = normalizeCardKey(elements.cardSearchName.value);
  const groupQuery = normalizeCardKey(elements.cardSearchGroup.value);
  const cycleQuery = normalizeCardKey(elements.cardSearchCycle.value);

  return cards.filter((card) => {
    const matchesName = !nameQuery
      || normalizeCardKey(card.cardName).includes(nameQuery)
      || normalizeCardKey(card.bankName).includes(nameQuery);
    const matchesGroup = !groupQuery || normalizeCardKey(card.sharedGroup).includes(groupQuery);
    const cycleText = `${card.statementDate || ""} ${formatOrdinalDay(card.statementDate)} ${card.dueDate || ""} ${formatOrdinalDay(card.dueDate)}`;
    const matchesCycle = !cycleQuery || normalizeCardKey(cycleText).includes(cycleQuery);
    return matchesName && matchesGroup && matchesCycle;
  });
}

function getPortfolioEffectiveLimitTotal(cards) {
  const seenSharedGroups = new Set();
  return cards.reduce((sum, card) => {
    const sharedGroup = getSharedLimitGroupByName(card.sharedGroup);
    if (!sharedGroup) {
      return sum + Number(card.creditLimit || 0);
    }
    const key = normalizeCardKey(sharedGroup.groupName);
    if (seenSharedGroups.has(key)) {
      return sum;
    }
    seenSharedGroups.add(key);
    return sum + Number(sharedGroup.totalLimit || 0);
  }, 0);
}

function renderCardsSearchSummary(filteredCards, totalCount) {
  const nameQuery = elements.cardSearchName.value.trim();
  const groupQuery = elements.cardSearchGroup.value.trim();
  const cycleQuery = elements.cardSearchCycle.value.trim();

  if (!nameQuery && !groupQuery && !cycleQuery) {
    elements.cardsSearchSummary.className = "result-card empty-state";
    elements.cardsSearchSummary.textContent = "Search cards by name, bank, group, or billing cycle to review your portfolio faster.";
    return;
  }

  const totalLimit = getPortfolioEffectiveLimitTotal(filteredCards);
  const availableLimit = filteredCards.reduce((sum, card) => sum + getCardAvailableLimit(card), 0);
  const outstanding = filteredCards.reduce((sum, card) => sum + getCardOutstanding(card), 0);

  elements.cardsSearchSummary.className = "result-card";
  elements.cardsSearchSummary.innerHTML = `
    <div class="row-between">
      <div>
        <h4>${filteredCards.length} matching card${filteredCards.length === 1 ? "" : "s"}</h4>
        <p class="mini-text">${[
          nameQuery ? `Card / Bank: ${escapeHtml(nameQuery)}` : "",
          groupQuery ? `Group: ${escapeHtml(groupQuery)}` : "",
          cycleQuery ? `Cycle: ${escapeHtml(cycleQuery)}` : ""
        ].filter(Boolean).join(" | ")}</p>
      </div>
      <span class="pill good">${totalCount} total</span>
    </div>
    <p><strong>Total limit:</strong> ${formatCurrency(totalLimit)}</p>
    <p><strong>Available limit:</strong> ${formatCurrency(availableLimit)}</p>
    <p><strong>Outstanding:</strong> ${formatCurrency(outstanding)}</p>
  `;
}

function clearCardSearch() {
  elements.cardSearchName.value = "";
  elements.cardSearchGroup.value = "";
  elements.cardSearchCycle.value = "";
  renderCardsTable();
}

function renderSharedLimitGroupsTable() {
  if (!state.sharedLimitGroups.length) {
    elements.sharedLimitGroupsTableBody.innerHTML = `<tr><td colspan="7">No shared limit groups added yet.</td></tr>`;
    return;
  }

  const groups = [...state.sharedLimitGroups].sort((a, b) => String(a.groupName || "").localeCompare(String(b.groupName || "")));
  elements.sharedLimitGroupsTableBody.innerHTML = groups.map((group) => {
    const members = getCardsInSharedGroup(group.groupName);
    const used = getSharedGroupNetUtilizedAmount(group.groupName);
    const available = Math.max(Number(group.totalLimit || 0) - used, 0);
    return `
      <tr>
        <td>${escapeHtml(group.groupName)}</td>
        <td>${formatCurrency(group.totalLimit)}</td>
        <td>${escapeHtml(members.map((card) => card.cardName).join(", ") || "-")}</td>
        <td>${formatCurrency(used)}</td>
        <td>${formatCurrency(available)}</td>
        <td>${escapeHtml(group.notes || "-")}</td>
        <td><button class="table-action-btn" type="button" data-edit-shared-group="${group.id}">Edit</button></td>
      </tr>
    `;
  }).join("");
}

function populateSharedLimitGroupMemberOptions() {
  if (!elements.sharedLimitGroupMembers) {
    return;
  }
  const selected = getMultiSelectValues(elements.sharedLimitGroupMembers);
  const cards = getSortedCards();
  elements.sharedLimitGroupMembers.innerHTML = cards.map((card) => `<option value="${escapeHtml(card.id)}">${escapeHtml(card.cardName)}</option>`).join("");
  setMultiSelectValues(elements.sharedLimitGroupMembers, selected);
}

function renderTransactionSuggestions() {
  const merchants = [...new Set(state.transactions.map((item) => String(item.merchant || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  const descriptions = [...new Set(state.transactions.map((item) => String(item.description || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  elements.merchantSuggestions.innerHTML = merchants.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  elements.descriptionSuggestions.innerHTML = descriptions.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
}
async function handleTransactionSubmit(event) {
  event.preventDefault();

  const card = getCardById(elements.transactionCard.value);
  if (!card) {
    showToast("Please add a card before saving a transaction.");
    return;
  }

  const amount = Number(elements.transactionAmount.value || 0);
  if (amount <= 0) {
    showToast("Please enter a valid amount.");
    return;
  }

  const transactionNature = normalizeTransactionNature(elements.transactionNature.value);
  const rewardPoints = calculateRewardPoints(card, amount, elements.transactionCategory.value, elements.transactionMerchant.value.trim(), transactionNature);
  const transaction = normalizeTransaction({
    id: state.ui.editingTransactionId || createId(),
    date: elements.transactionDate.value,
    cardId: card.id,
    cardName: card.cardName,
    amount,
    category: elements.transactionCategory.value,
    description: elements.transactionDescription.value.trim(),
    merchant: elements.transactionMerchant.value.trim(),
    transactionNature,
    transactionType: elements.transactionType.value,
    paymentStatus: elements.paymentStatus.value,
    billPdfLink: elements.billPdfLink.value.trim(),
    unbilledAmount: Number(elements.unbilledAmount.value || (isExpenseNature(transactionNature) ? amount : 0)),
    notes: elements.transactionNotes.value.trim(),
    rewardPoints
  });

  if (state.ui.editingTransactionId) {
    const index = state.transactions.findIndex((item) => item.id === state.ui.editingTransactionId);
    if (index >= 0) {
      state.transactions[index] = transaction;
    }
  } else {
    state.transactions.unshift(transaction);
  }

  saveLocalState();
  renderApp();
  const editing = Boolean(state.ui.editingTransactionId);
  resetTransactionForm();
  closeTransactionModal();
  showToast(editing ? "Transaction updated successfully." : "Transaction saved successfully.");
  await syncRecord("transactions", transaction);
}

async function handleCardSubmit(event) {
  event.preventDefault();

  try {
    const parsedRewardRules = parseRewardRulesInput(elements.rewardRules.value);
    const parsedAdvancedRules = parseAdvancedRewardRulesInput(elements.advancedRewardRules.value);
    const parsedMilestones = parseMilestonesInput(elements.milestones.value);
    const parsedAnnualBenefits = parseAnnualBenefitsInput(elements.annualBenefits.value);

    const card = normalizeCard({
      id: state.ui.editingCardId || createId(),
      cardName: elements.cardName.value.trim(),
      bankName: elements.bankName.value.trim(),
      creditLimit: Number(elements.creditLimit.value || 0),
      billingCycleDate: Number(elements.billingCycleDate.value || 0),
      dueDate: Number(elements.dueDate.value || 0),
      statementDate: Number(elements.statementDate.value || 0),
      anniversaryDate: elements.anniversaryDate.value,
      renewalTarget: Number(elements.renewalTarget.value || 0),
      sharedGroup: elements.sharedGroup.value.trim(),
      preferredCategories: splitCsv(elements.preferredCategories.value),
      preferredTransactionTypes: splitCsv(elements.preferredTransactionTypes.value),
      supportedTransactionNature: splitCsv(elements.supportedTransactionNature.value),
      rewardRules: parsedRewardRules.rewardRules,
      rewardRulesInput: parsedRewardRules.rawInput,
      advancedRewardRules: parsedAdvancedRules,
      advancedRewardRulesInput: elements.advancedRewardRules.value.trim(),
      milestones: parsedMilestones,
      milestonesInput: elements.milestones.value.trim(),
      annualBenefits: parsedAnnualBenefits,
      annualBenefitsInput: elements.annualBenefits.value.trim()
    });

    if (state.ui.editingCardId) {
      const index = state.cards.findIndex((item) => item.id === state.ui.editingCardId);
      if (index >= 0) {
        state.cards[index] = card;
      }
    } else {
      state.cards.push(card);
    }

    recalculateRewardsForCard(card);

    saveLocalState();
    renderApp();
    const editing = Boolean(state.ui.editingCardId);
    resetCardForm();
    showToast(editing ? "Card updated successfully." : "Card saved successfully.");
    await syncRecord("cards", card);
  } catch (error) {
    showToast(error.message || "Unable to save card.");
  }
}

async function handleSharedLimitGroupSubmit(event) {
  event.preventDefault();

  const group = normalizeSharedLimitGroup({
    id: state.ui.editingSharedLimitGroupId || createId(),
    groupName: elements.sharedLimitGroupName.value.trim(),
    totalLimit: Number(elements.sharedLimitGroupLimit.value || 0),
    memberCardIds: getMultiSelectValues(elements.sharedLimitGroupMembers),
    notes: elements.sharedLimitGroupNotes.value.trim()
  });

  if (!group.groupName) {
    showToast("Please enter a valid shared group name.");
    return;
  }

  if (state.ui.editingSharedLimitGroupId) {
    const index = state.sharedLimitGroups.findIndex((item) => item.id === state.ui.editingSharedLimitGroupId);
    if (index >= 0) {
      state.sharedLimitGroups[index] = group;
    }
  } else {
    state.sharedLimitGroups.push(group);
  }

  const editing = Boolean(state.ui.editingSharedLimitGroupId);
  saveLocalState();
  renderApp();
  resetSharedLimitGroupForm();
  showToast(editing ? "Shared limit group updated." : "Shared limit group saved.");
  await syncRecord("sharedLimitGroups", group);
}

async function handleRewardRuleSubmit(event) {
  event.preventDefault();

  const card = getCardById(elements.rewardRuleCard.value);
  if (!card) {
    showToast("Please select a valid card.");
    return;
  }

  const rewardRule = normalizeAdvancedRewardRule({
    id: state.ui.editingRewardRuleId || createId(),
    cardName: card.cardName,
    category: elements.rewardRuleCategory.value.trim(),
    scope: elements.rewardRuleScope.value,
    merchant: resolveRewardRuleMerchantValue(elements.rewardRuleScope.value, elements.rewardRuleMerchant.value.trim()),
    baseUnit: Number(elements.rewardRuleBaseUnit.value || 0),
    rewardPerUnit: Number(elements.rewardRuleRewardPerUnit.value || 0),
    multiplier: Number(elements.rewardRuleMultiplier.value || 1),
    priority: Number(elements.rewardRulePriority.value || 999)
  });

  if (state.ui.editingRewardRuleId) {
    const index = state.rewardRulesCatalog.findIndex((item) => item.id === state.ui.editingRewardRuleId);
    if (index >= 0) {
      state.rewardRulesCatalog[index] = rewardRule;
    }
  } else {
    state.rewardRulesCatalog.push(rewardRule);
  }

  recalculateRewardsForCard(card);
  const wasEditing = Boolean(state.ui.editingRewardRuleId);
  saveLocalState();
  renderApp();
  resetRewardRuleForm();
  showToast(wasEditing ? "Reward rule updated." : "Reward rule saved.");
  await syncRecord("rewardRules", rewardRule);
}

async function handleMilestoneSubmit(event) {
  event.preventDefault();

  const card = getCardById(elements.milestoneCard.value);
  if (!card) {
    showToast("Please select a valid card.");
    return;
  }

  const milestone = normalizeMilestoneRule({
    id: state.ui.editingMilestoneId || createId(),
    cardName: card.cardName,
    milestoneType: elements.milestoneType.value.trim(),
    cycleType: elements.milestoneCycleType.value,
    target: Number(elements.milestoneTarget.value || 0),
    rewardType: elements.milestoneRewardType.value.trim(),
    rewardValue: Number(elements.milestoneRewardValue.value || 0),
    notes: elements.milestoneNotes.value.trim()
  });

  if (state.ui.editingMilestoneId) {
    const index = state.milestoneCatalog.findIndex((item) => item.id === state.ui.editingMilestoneId);
    if (index >= 0) {
      state.milestoneCatalog[index] = milestone;
    }
  } else {
    state.milestoneCatalog.push(milestone);
  }

  const wasEditing = Boolean(state.ui.editingMilestoneId);
  saveLocalState();
  renderApp();
  resetMilestoneForm();
  showToast(wasEditing ? "Milestone updated." : "Milestone saved.");
  await syncRecord("milestones", milestone);
}

function handleSettingsSubmit(event) {
  event.preventDefault();
  state.settings.scriptUrl = elements.scriptUrl.value.trim();
  localStorage.setItem(SCRIPT_URL_KEY, state.settings.scriptUrl);
  renderConnectionStatus();
  showToast("Settings saved.");
}

function handleCardTableActions(event) {
  const openButton = event.target.closest("[data-open-card-row]");
  if (openButton) {
    openCardAccount(openButton.getAttribute("data-open-card-row"));
    return;
  }
  const editButton = event.target.closest("[data-edit-card]");
  if (!editButton) {
    return;
  }
  startCardEdit(editButton.getAttribute("data-edit-card"));
}

function handleSharedLimitGroupTableActions(event) {
  const editButton = event.target.closest("[data-edit-shared-group]");
  if (!editButton) {
    return;
  }
  startSharedLimitGroupEdit(editButton.getAttribute("data-edit-shared-group"));
}

function handleTransactionTableActions(event) {
  const button = event.target.closest("[data-edit-transaction]");
  if (!button) {
    return;
  }
  startTransactionEdit(button.getAttribute("data-edit-transaction"));
}

function handleRewardRuleTableActions(event) {
  const button = event.target.closest("[data-edit-reward-rule]");
  if (!button) {
    return;
  }
  startRewardRuleEdit(button.getAttribute("data-edit-reward-rule"));
}

function handleMilestoneTableActions(event) {
  const button = event.target.closest("[data-edit-milestone]");
  if (!button) {
    return;
  }
  startMilestoneEdit(button.getAttribute("data-edit-milestone"));
}

function handleOpenCardRequest(event) {
  const target = event.target.closest("[data-open-card]");
  if (!target) {
    return;
  }
  openCardAccount(target.getAttribute("data-open-card"));
}

function openCardAccount(cardId) {
  if (!getCardById(cardId)) {
    return;
  }
  state.ui.activeCardId = cardId;
  if (isMobileViewport()) {
    closeMobileNav();
  }
  activateView("account-view");
  renderApp();
}

function startCardEdit(cardId) {
  const card = getCardById(cardId);
  if (!card) {
    return;
  }

  state.ui.editingCardId = cardId;
  elements.cardName.value = card.cardName || "";
  elements.bankName.value = card.bankName || "";
  elements.creditLimit.value = card.creditLimit || "";
  elements.billingCycleDate.value = card.billingCycleDate || "";
  elements.dueDate.value = card.dueDate || "";
  elements.statementDate.value = card.statementDate || "";
  elements.anniversaryDate.value = card.anniversaryDate || "";
  elements.renewalTarget.value = card.renewalTarget || "";
  elements.sharedGroup.value = card.sharedGroup || "";
  elements.preferredCategories.value = (card.preferredCategories || []).join(", ");
  elements.preferredTransactionTypes.value = (card.preferredTransactionTypes || []).join(", ");
  elements.supportedTransactionNature.value = (card.supportedTransactionNature || []).join(", ");
  elements.rewardRules.value = card.rewardRulesInput || formatRewardRulesForEditor(card.rewardRules);
  elements.advancedRewardRules.value = card.advancedRewardRulesInput || formatJsonForEditor(card.advancedRewardRules);
  elements.milestones.value = card.milestonesInput || formatJsonForEditor(card.milestones);
  elements.annualBenefits.value = card.annualBenefitsInput || formatJsonForEditor(card.annualBenefits);
  elements.cancelCardEditButton.hidden = false;
  activateView("cards-view");
  elements.cardName.focus();
}

function startTransactionEdit(transactionId) {
  const transaction = state.transactions.find((item) => item.id === transactionId);
  if (!transaction) {
    return;
  }

  state.ui.editingTransactionId = transactionId;
  elements.transactionModalTitle.textContent = "Edit Transaction";
  elements.transactionDate.value = transaction.date || "";
  elements.transactionCard.value = transaction.cardId || "";
  elements.transactionAmount.value = transaction.amount || "";
  elements.transactionCategory.value = transaction.category || CATEGORIES[0];
  elements.transactionDescription.value = transaction.description || "";
  elements.transactionMerchant.value = transaction.merchant || "";
  elements.transactionNature.value = normalizeTransactionNature(transaction.transactionNature);
  elements.transactionType.value = transaction.transactionType || TRANSACTION_TYPES[0];
  elements.paymentStatus.value = transaction.paymentStatus || PAYMENT_STATUSES[0];
  elements.billPdfLink.value = transaction.billPdfLink || "";
  elements.unbilledAmount.value = transaction.unbilledAmount || "";
  elements.transactionNotes.value = transaction.notes || "";
  elements.cancelTransactionEditButton.hidden = false;
  updateTransactionPreview();
  openTransactionModal();
}

function startRewardRuleEdit(ruleId) {
  const rule = state.rewardRulesCatalog.find((item) => item.id === ruleId);
  if (!rule) {
    return;
  }

  state.ui.editingRewardRuleId = ruleId;
  const matchingCard = state.cards.find((item) => normalizeCardKey(item.cardName) === normalizeCardKey(rule.cardName));
  if (matchingCard) {
    elements.rewardRuleCard.value = matchingCard.id;
  }
  elements.rewardRuleCategory.value = rule.category || "";
  elements.rewardRuleScope.value = rule.scope || inferRewardRuleScope(rule);
  elements.rewardRuleMerchant.value = normalizeCardKey(rule.merchant) === "any" ? "" : (rule.merchant || "");
  elements.rewardRuleBaseUnit.value = rule.baseUnit || "";
  elements.rewardRuleRewardPerUnit.value = rule.rewardPerUnit || "";
  elements.rewardRuleMultiplier.value = rule.multiplier || 1;
  elements.rewardRulePriority.value = rule.priority || 999;
  elements.cancelRewardRuleEditButton.hidden = false;
  updateRewardRuleScopeUI();
  activateView("reward-rules-view");
}

function startMilestoneEdit(ruleId) {
  const rule = state.milestoneCatalog.find((item) => item.id === ruleId);
  if (!rule) {
    return;
  }

  state.ui.editingMilestoneId = ruleId;
  const matchingCard = state.cards.find((item) => normalizeCardKey(item.cardName) === normalizeCardKey(rule.cardName));
  if (matchingCard) {
    elements.milestoneCard.value = matchingCard.id;
  }
  elements.milestoneType.value = rule.milestoneType || "";
  elements.milestoneCycleType.value = rule.cycleType || "Annual";
  elements.milestoneTarget.value = rule.target || "";
  elements.milestoneRewardType.value = rule.rewardType || "";
  elements.milestoneRewardValue.value = rule.rewardValue || "";
  elements.milestoneNotes.value = rule.notes || "";
  elements.cancelMilestoneEditButton.hidden = false;
  activateView("milestones-view");
}

function startSharedLimitGroupEdit(groupId) {
  const group = state.sharedLimitGroups.find((item) => item.id === groupId);
  if (!group) {
    return;
  }

  const fallbackMemberIds = (!group.memberCardIds || !group.memberCardIds.length)
    ? state.cards.filter((card) => normalizeCardKey(card.sharedGroup) === normalizeCardKey(group.groupName)).map((card) => card.id)
    : group.memberCardIds;

  state.ui.editingSharedLimitGroupId = groupId;
  elements.sharedLimitGroupName.value = group.groupName || "";
  elements.sharedLimitGroupLimit.value = group.totalLimit || "";
  setMultiSelectValues(elements.sharedLimitGroupMembers, fallbackMemberIds);
  elements.sharedLimitGroupNotes.value = group.notes || "";
  elements.cancelSharedLimitGroupEditButton.hidden = false;
  activateView("cards-view");
}

function resetCardForm() {
  state.ui.editingCardId = null;
  elements.cardForm.reset();
  elements.cancelCardEditButton.hidden = true;
  elements.anniversaryDate.value = new Date().toISOString().split("T")[0];
}

function resetSharedLimitGroupForm() {
  state.ui.editingSharedLimitGroupId = null;
  elements.sharedLimitGroupForm.reset();
  setMultiSelectValues(elements.sharedLimitGroupMembers, []);
  elements.cancelSharedLimitGroupEditButton.hidden = true;
}

function resetTransactionForm() {
  state.ui.editingTransactionId = null;
  elements.transactionForm.reset();
  elements.cancelTransactionEditButton.hidden = true;
  elements.transactionModalTitle.textContent = "New Transaction";
  setDefaultDates();
  populateCardDropdowns();
  updateTransactionPreview();
}

function resetRewardRuleForm() {
  state.ui.editingRewardRuleId = null;
  elements.rewardRuleForm.reset();
  elements.rewardRuleScope.value = "category";
  elements.cancelRewardRuleEditButton.hidden = true;
  populateCardDropdowns();
  updateRewardRuleScopeUI();
}

function resetMilestoneForm() {
  state.ui.editingMilestoneId = null;
  elements.milestoneForm.reset();
  elements.cancelMilestoneEditButton.hidden = true;
  populateCardDropdowns();
}

function openTransactionModal() {
  if (state.ui.activeCardId && getCardById(state.ui.activeCardId) && !state.ui.editingTransactionId) {
    elements.transactionCard.value = state.ui.activeCardId;
  }
  elements.transactionModal.classList.remove("hidden");
}

function closeTransactionModal() {
  elements.transactionModal.classList.add("hidden");
}

async function syncFromRemote() {
  if (!state.settings.scriptUrl) {
    showToast("Add your Apps Script URL first.");
    return;
  }

  try {
    const response = await fetch(`${state.settings.scriptUrl}?action=fetchAll`, { method: "GET" });
    const payload = await response.json();
    if (!payload.success) {
      throw new Error(payload.message || "Unable to sync.");
    }

    const remoteCards = Array.isArray(payload.data.cards) ? payload.data.cards.map(normalizeCard) : [];
    state.cards = mergeRecordsById(removeStarterCardsFromState(state.cards, remoteCards), remoteCards, normalizeCard);
    state.transactions = mergeRecordsById(state.transactions, Array.isArray(payload.data.transactions) ? payload.data.transactions.map(normalizeTransaction) : [], normalizeTransaction);
    state.rewardActivities = mergeRecordsById(state.rewardActivities, Array.isArray(payload.data.rewardActivities) ? payload.data.rewardActivities.map(normalizeRewardActivity) : [], normalizeRewardActivity);
    state.rewardRulesCatalog = mergeRecordsById(state.rewardRulesCatalog, Array.isArray(payload.data.rewardRules) ? payload.data.rewardRules.map(normalizeAdvancedRewardRule) : [], normalizeAdvancedRewardRule);
    state.milestoneCatalog = mergeRecordsById(state.milestoneCatalog, Array.isArray(payload.data.milestones) ? payload.data.milestones.map(normalizeMilestoneRule) : [], normalizeMilestoneRule);
    state.sharedLimitGroups = mergeRecordsById(state.sharedLimitGroups, Array.isArray(payload.data.sharedLimitGroups) ? payload.data.sharedLimitGroups.map(normalizeSharedLimitGroup) : [], normalizeSharedLimitGroup);
    dedupeState();
    recalculateAllRewardPoints();
    state.lastSync = new Date().toISOString();

    saveLocalState();
    renderApp();
    showToast("Data synced from Google Sheets.");
  } catch (error) {
    console.error(error);
    showToast("Sync failed. Check your Apps Script deployment.");
  }
}

async function replaceLocalWithRemote() {
  if (!state.settings.scriptUrl) {
    showToast("Add your Apps Script URL first.");
    return;
  }

  try {
    const response = await fetch(`${state.settings.scriptUrl}?action=fetchAll`, { method: "GET" });
    const payload = await response.json();
    if (!payload.success) {
      throw new Error(payload.message || "Unable to sync.");
    }

    state.cards = Array.isArray(payload.data.cards) ? payload.data.cards.map(normalizeCard) : [];
    state.transactions = Array.isArray(payload.data.transactions) ? payload.data.transactions.map(normalizeTransaction) : [];
    state.rewardActivities = Array.isArray(payload.data.rewardActivities) ? payload.data.rewardActivities.map(normalizeRewardActivity) : [];
    state.rewardRulesCatalog = Array.isArray(payload.data.rewardRules) ? payload.data.rewardRules.map(normalizeAdvancedRewardRule) : [];
    state.milestoneCatalog = Array.isArray(payload.data.milestones) ? payload.data.milestones.map(normalizeMilestoneRule) : [];
    state.sharedLimitGroups = Array.isArray(payload.data.sharedLimitGroups) ? payload.data.sharedLimitGroups.map(normalizeSharedLimitGroup) : [];
    dedupeState();
    recalculateAllRewardPoints();
    state.lastSync = new Date().toISOString();

    saveLocalState();
    renderApp();
    showToast("Local data replaced with Google Sheets successfully.");
  } catch (error) {
    console.error(error);
    showToast("Replace failed. Check your Apps Script deployment.");
  }
}

function mergeRecordsById(localRecords, remoteRecords, normalizeFn) {
  const merged = new Map();

  (localRecords || []).forEach((record) => {
    const normalized = normalizeFn(record);
    merged.set(String(normalized.id), normalized);
  });

  (remoteRecords || []).forEach((record) => {
    const normalized = normalizeFn(record);
    merged.set(String(normalized.id), normalized);
  });

  return Array.from(merged.values());
}

async function syncRecord(sheetType, record) {
  if (!state.settings.scriptUrl) {
    return;
  }

  try {
    const response = await fetch(state.settings.scriptUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "upsertRecord",
        sheetType,
        record
      })
    });
    const payload = await response.json();
    if (!payload.success) {
      throw new Error(payload.message || "Unable to save record.");
    }

    state.lastSync = new Date().toISOString();
    saveLocalState();
    renderConnectionStatus();
  } catch (error) {
    console.error(error);
    showToast("Saved locally, but Google Sheets sync failed.");
  }
}

async function pushLocalDataToGoogleSheets() {
  if (!state.settings.scriptUrl) {
    showToast("Add your Apps Script URL first.");
    return;
  }

  const datasets = [
    { sheetType: "cards", records: state.cards },
    { sheetType: "transactions", records: state.transactions },
    { sheetType: "rewardActivities", records: state.rewardActivities },
    { sheetType: "rewardRules", records: state.rewardRulesCatalog },
    { sheetType: "milestones", records: state.milestoneCatalog },
    { sheetType: "sharedLimitGroups", records: state.sharedLimitGroups }
  ];

  const totalRecords = datasets.reduce((sum, item) => sum + item.records.length, 0);
  if (!totalRecords) {
    showToast("No local data found to push.");
    return;
  }

  elements.pushLocalDataButton.disabled = true;
  const originalLabel = elements.pushLocalDataButton.textContent;
  elements.pushLocalDataButton.textContent = "Pushing...";

  try {
    let pushed = 0;

    for (const dataset of datasets) {
      for (const record of dataset.records) {
        const response = await fetch(state.settings.scriptUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "upsertRecord",
            sheetType: dataset.sheetType,
            record
          })
        });

        const payload = await response.json();
        if (!payload.success) {
          throw new Error(payload.message || `Unable to upload ${dataset.sheetType}.`);
        }

        pushed += 1;
        elements.pushLocalDataButton.textContent = `Pushing ${pushed}/${totalRecords}`;
      }
    }

    state.lastSync = new Date().toISOString();
    saveLocalState();
    renderConnectionStatus();
    showToast(`Local data pushed successfully. ${pushed} record${pushed === 1 ? "" : "s"} uploaded.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "Bulk upload failed. Please try again.");
  } finally {
    elements.pushLocalDataButton.disabled = false;
    elements.pushLocalDataButton.textContent = originalLabel;
  }
}

function exportBackup() {
  const payload = {
    exportedAt: new Date().toISOString(),
    version: 1,
    data: {
      cards: state.cards,
      transactions: state.transactions,
      rewardActivities: state.rewardActivities,
      rewardRulesCatalog: state.rewardRulesCatalog,
      milestoneCatalog: state.milestoneCatalog,
      sharedLimitGroups: state.sharedLimitGroups,
      lastSync: state.lastSync
    }
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `credit-card-backup-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  showToast("Backup exported successfully.");
}

function importBackup(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || "{}"));
      const data = parsed.data || parsed;
      state.cards = Array.isArray(data.cards) ? data.cards.map(normalizeCard) : [];
      state.transactions = Array.isArray(data.transactions) ? data.transactions.map(normalizeTransaction) : [];
      state.rewardActivities = Array.isArray(data.rewardActivities) ? data.rewardActivities.map(normalizeRewardActivity) : [];
      state.rewardRulesCatalog = Array.isArray(data.rewardRulesCatalog) ? data.rewardRulesCatalog.map(normalizeAdvancedRewardRule) : [];
      state.milestoneCatalog = Array.isArray(data.milestoneCatalog) ? data.milestoneCatalog.map(normalizeMilestoneRule) : [];
      state.sharedLimitGroups = Array.isArray(data.sharedLimitGroups) ? data.sharedLimitGroups.map(normalizeSharedLimitGroup) : [];
      recalculateAllRewardPoints();
      state.lastSync = data.lastSync || null;
      if (!getCardById(state.ui.activeCardId)) {
        state.ui.activeCardId = null;
      }
      saveLocalState();
      renderApp();
      resetCardForm();
      resetTransactionForm();
      showToast("Backup imported successfully.");
    } catch (error) {
      console.error(error);
      showToast("Backup import failed. Please use a valid backup file.");
    } finally {
      elements.backupFileInput.value = "";
    }
  };
  reader.readAsText(file);
}

function updateTransactionPreview() {
  const card = getCardById(elements.transactionCard.value);
  const amount = Number(elements.transactionAmount.value || 0);
  const category = elements.transactionCategory.value;
  const merchant = elements.transactionMerchant.value.trim();
  const nature = normalizeTransactionNature(elements.transactionNature.value);

  if (!card || amount <= 0) {
    elements.transactionPreview.textContent = "Reward preview will appear here.";
    return;
  }

  if (!isRewardEligibleNature(nature)) {
    elements.transactionPreview.textContent = "This transaction nature does not earn reward points.";
    return;
  }

  const points = calculateRewardPoints(card, amount, category, merchant, nature);
  elements.transactionPreview.textContent = `Expected reward for ${card.cardName}: ${formatPoints(points)}.`;
}

function renderConnectionStatus() {
  if (state.settings.scriptUrl) {
    elements.connectionStatus.textContent = "Ready for sync";
    elements.lastSyncText.textContent = state.lastSync
      ? `Last sync: ${formatDateTime(state.lastSync)}`
      : "Google Sheets URL saved. Manual sync is available.";
  } else {
    elements.connectionStatus.textContent = "Local only";
    elements.lastSyncText.textContent = "Add your Apps Script URL in Settings to enable Google Sheets sync.";
  }
}
function getSortedCards() {
  return [...state.cards].sort((a, b) => a.cardName.localeCompare(b.cardName));
}

function recalculateRewardsForCard(card) {
  state.transactions = state.transactions.map((transaction) => {
    if (!transactionMatchesCard(transaction, card)) {
      return transaction;
    }

    return {
      ...transaction,
      rewardPoints: calculateRewardPoints(
        card,
        Number(transaction.amount || 0),
        transaction.category,
        transaction.merchant,
        transaction.transactionNature
      )
    };
  });
}

function recalculateAllRewardPoints() {
  state.transactions = state.transactions.map((transaction) => {
    const card = getCardById(transaction.cardId) || state.cards.find((item) => normalizeCardKey(item.cardName) === normalizeCardKey(transaction.cardName));
    if (!card) {
      return transaction;
    }

    return {
      ...transaction,
      rewardPoints: calculateRewardPoints(
        card,
        Number(transaction.amount || 0),
        transaction.category,
        transaction.merchant,
        transaction.transactionNature
      )
    };
  });
}

function getSortedTransactions() {
  return [...state.transactions].sort((a, b) => new Date(b.date) - new Date(a.date));
}

function getAllTransactionFilters() {
  const merchantQuery = normalizeCardKey(elements.transactionSearchMerchant.value);
  const descriptionQuery = normalizeCardKey(elements.transactionSearchDescription.value);
  const cardId = elements.transactionSearchCard.value;
  const category = elements.transactionSearchCategory.value.trim();
  const nature = elements.transactionSearchNature.value.trim();
  const fromDate = elements.transactionSearchFromDate.value;
  const toDate = elements.transactionSearchToDate.value;
  const minAmount = Number(elements.transactionSearchMinAmount.value || 0);
  const maxAmount = elements.transactionSearchMaxAmount.value ? Number(elements.transactionSearchMaxAmount.value) : null;

  return [
    (item) => !merchantQuery || normalizeCardKey(item.merchant).includes(merchantQuery),
    (item) => !descriptionQuery || normalizeCardKey(item.description).includes(descriptionQuery),
    (item) => !cardId || transactionMatchesCard(item, getCardById(cardId)),
    (item) => !category || String(item.category || "") === category,
    (item) => !nature || normalizeTransactionNature(item.transactionNature) === normalizeTransactionNature(nature),
    (item) => !fromDate || startOfDay(new Date(item.date)) >= startOfDay(new Date(fromDate)),
    (item) => !toDate || startOfDay(new Date(item.date)) <= startOfDay(new Date(toDate)),
    (item) => Number(item.amount || 0) >= minAmount,
    (item) => maxAmount == null || Number(item.amount || 0) <= maxAmount
  ];
}

function getFilteredAccountTransactions(card) {
  const merchantQuery = normalizeCardKey(elements.accountSearchMerchant.value);
  const descriptionQuery = normalizeCardKey(elements.accountSearchDescription.value);

  return getCardTransactions(card).filter((item) => {
    const merchantMatch = !merchantQuery || normalizeCardKey(item.merchant).includes(merchantQuery);
    const descriptionMatch = !descriptionQuery || normalizeCardKey(item.description).includes(descriptionQuery);
    return merchantMatch && descriptionMatch;
  });
}

function getCardTransactions(card) {
  return getSortedTransactions().filter((transaction) => transactionMatchesCard(transaction, card));
}

function getCardExpenseTransactions(card) {
  return getCardTransactions(card).filter((item) => isExpenseNature(item.transactionNature));
}

function getCardCreditTransactions(card) {
  return getCardTransactions(card).filter((item) => isCreditNature(item.transactionNature));
}

function getSharedLimitGroupByName(groupName) {
  const key = normalizeCardKey(groupName);
  if (!key) {
    return null;
  }
  return state.sharedLimitGroups.find((group) => normalizeCardKey(group.groupName) === key) || null;
}

function getSharedLimitGroupForCard(card) {
  const group = getSharedLimitGroupByName(card.sharedGroup);
  if (!group) {
    return null;
  }
  const memberIds = Array.isArray(group.memberCardIds) ? group.memberCardIds.map(String) : [];
  if (!memberIds.includes(String(card.id))) {
    return null;
  }
  return group;
}

function getCardsInSharedGroup(groupName) {
  const group = getSharedLimitGroupByName(groupName);
  if (!group) {
    return [];
  }
  const memberIds = Array.isArray(group.memberCardIds) ? group.memberCardIds.map(String) : [];
  if (!memberIds.length) {
    return [];
  }
  return state.cards.filter((card) => memberIds.includes(String(card.id)));
}

function getCardNetUtilizedAmount(card) {
  const debits = getCardExpenseTransactions(card).reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const credits = getCardCreditTransactions(card).reduce((sum, item) => sum + Number(item.amount || 0), 0);
  return Math.max(debits - credits, 0);
}

function getEffectiveCardLimit(card) {
  const sharedGroup = getSharedLimitGroupForCard(card);
  return sharedGroup ? Number(sharedGroup.totalLimit || 0) : Number(card.creditLimit || 0);
}

function getSharedGroupNetUtilizedAmount(groupName) {
  return getCardsInSharedGroup(groupName).reduce((sum, card) => sum + getCardNetUtilizedAmount(card), 0);
}

function getCardAvailableLimit(card) {
  const sharedGroup = getSharedLimitGroupForCard(card);
  if (sharedGroup) {
    return Math.max(Number(sharedGroup.totalLimit || 0) - getSharedGroupNetUtilizedAmount(sharedGroup.groupName), 0);
  }
  return Math.max(Number(card.creditLimit || 0) - getCardNetUtilizedAmount(card), 0);
}

function getCardOutstanding(card) {
  const unpaidDebits = getCardExpenseTransactions(card)
    .filter((item) => normalizePaymentStatus(item.paymentStatus) !== "Paid")
    .reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const credits = getCardCreditTransactions(card).reduce((sum, item) => sum + Number(item.amount || 0), 0);
  return Math.max(unpaidDebits - credits, 0);
}

function getCardUnbilledAmount(card) {
  const expenses = getCardExpenseTransactions(card).reduce((sum, item) => sum + Number(item.unbilledAmount || 0), 0);
  const credits = getCardCreditTransactions(card).reduce((sum, item) => sum + Number(item.amount || 0), 0);
  return Math.max(expenses - credits, 0);
}

function getCardAnniversarySpend(card) {
  const startDate = getAnniversaryCycleStart(card);
  const expenses = getCardExpenseTransactions(card).filter((item) => new Date(item.date) >= startDate);
  return expenses.reduce((sum, item) => sum + Number(item.amount || 0), 0);
}

function getCardRewardEarnedFromTransactions(card) {
  return getCardExpenseTransactions(card).reduce((sum, item) => {
    return sum + calculateRewardPoints(
      card,
      Number(item.amount || 0),
      item.category,
      item.merchant,
      item.transactionNature
    );
  }, 0);
}

function getCardRedeemedPoints(card) {
  return state.rewardActivities
    .filter((item) => rewardActivityMatchesCard(item, card) && normalizeCardKey(item.action) === "redeemed")
    .reduce((sum, item) => sum + Number(item.points || 0), 0);
}

function getCardExpiredPoints(card) {
  return state.rewardActivities
    .filter((item) => rewardActivityMatchesCard(item, card) && normalizeCardKey(item.action) === "expiry")
    .reduce((sum, item) => sum + Number(item.points || 0), 0);
}

function getCardRewardTotal(card) {
  return Math.max(getCardRewardEarnedFromTransactions(card) - getCardRedeemedPoints(card) - getCardExpiredPoints(card), 0);
}

function getAlertData() {
  return getSortedCards().map((card) => {
    const liveDueDate = getLiveDueDate(card);
    const daysLeft = diffDays(startOfDay(liveDueDate), startOfDay(new Date()));
    return {
      cardId: card.id,
      cardName: card.cardName,
      liveDueDate,
      outstanding: getCardOutstanding(card),
      daysLeft
    };
  }).sort((a, b) => a.daysLeft - b.daysLeft);
}

function formatOrdinalDay(value) {
  const day = Number(value || 0);
  if (!day) {
    return "-";
  }
  const mod10 = day % 10;
  const mod100 = day % 100;
  let suffix = "th";
  if (mod10 === 1 && mod100 !== 11) suffix = "st";
  if (mod10 === 2 && mod100 !== 12) suffix = "nd";
  if (mod10 === 3 && mod100 !== 13) suffix = "rd";
  return `${day}${suffix}`;
}

function getLatestBillCycle(card) {
  const today = startOfDay(new Date());
  const statementDay = Number(card.statementDate || card.billingCycleDate || 1);
  const currentStatementDate = getLatestStatementDate(today, statementDay);
  const previousStatementDate = createDateForDay(currentStatementDate.getFullYear(), currentStatementDate.getMonth() - 1, statementDay);
  const liveDueDate = createDateForDay(currentStatementDate.getFullYear(), currentStatementDate.getMonth(), Number(card.dueDate || 1));
  const paymentWindowEnd = today < liveDueDate ? today : liveDueDate;

  const transactions = getCardTransactions(card);
  const billedAmount = transactions
    .filter((item) => isExpenseNature(item.transactionNature))
    .filter((item) => {
      const txDate = startOfDay(new Date(item.date));
      return txDate > previousStatementDate && txDate <= currentStatementDate;
    })
    .reduce((sum, item) => sum + Number(item.amount || 0), 0);

  const paidAmount = transactions
    .filter((item) => isCreditNature(item.transactionNature))
    .filter((item) => {
      const txDate = startOfDay(new Date(item.date));
      return txDate > currentStatementDate && txDate <= paymentWindowEnd;
    })
    .reduce((sum, item) => sum + Number(item.amount || 0), 0);

  const cycleMonthDate = new Date(previousStatementDate);
  cycleMonthDate.setDate(cycleMonthDate.getDate() + 1);

  return {
    label: new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric" }).format(cycleMonthDate),
    billedAmount,
    paidAmount,
    pendingAmount: Math.max(billedAmount - paidAmount, 0),
    statementDate: currentStatementDate,
    dueDate: liveDueDate
  };
}

function getLiveDueDate(card) {
  const now = new Date();
  const currentMonthDue = createDateForDay(now.getFullYear(), now.getMonth(), Number(card.dueDate || 1));
  if (currentMonthDue >= startOfDay(now)) {
    return currentMonthDue;
  }
  return createDateForDay(now.getFullYear(), now.getMonth() + 1, Number(card.dueDate || 1));
}

function getLatestStatementDate(referenceDate, statementDay) {
  const currentMonthStatement = createDateForDay(referenceDate.getFullYear(), referenceDate.getMonth(), statementDay);
  if (currentMonthStatement <= referenceDate) {
    return currentMonthStatement;
  }
  return createDateForDay(referenceDate.getFullYear(), referenceDate.getMonth() - 1, statementDay);
}

function getAnniversaryCycleStart(card) {
  const today = new Date();
  const anniversary = card.anniversaryDate ? new Date(card.anniversaryDate) : today;
  const cycleThisYear = new Date(today.getFullYear(), anniversary.getMonth(), anniversary.getDate());
  if (startOfDay(today) >= startOfDay(cycleThisYear)) {
    return cycleThisYear;
  }
  return new Date(today.getFullYear() - 1, anniversary.getMonth(), anniversary.getDate());
}

function getMilestoneStatus(card) {
  const milestones = getMilestonesForCard(card);
  if (!milestones.length) {
    return "No milestone configured.";
  }

  const groupedByCycle = milestones.reduce((groups, item) => {
    const key = item.cycleType || "Annual";
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(item);
    return groups;
  }, {});

  const cycleMessages = Object.keys(groupedByCycle).sort().map((cycleType) => {
    const spend = getCycleSpendForMilestone(card, cycleType);
    const sorted = [...groupedByCycle[cycleType]].sort((a, b) => Number(a.target || 0) - Number(b.target || 0));
    const nextMilestone = sorted.find((item) => Number(item.target || 0) > spend);
    if (!nextMilestone) {
      return `${cycleType}: highest milestone achieved`;
    }

    const remaining = Math.max(Number(nextMilestone.target || 0) - spend, 0);
    const label = nextMilestone.reward || nextMilestone.benefit || nextMilestone.notes || nextMilestone.rewardType || "Next milestone";
    return `${cycleType}: ${label} in ${formatCurrency(remaining)}`;
  });

  return cycleMessages.join(" | ");
}

function calculateRewardPoints(card, amount, category, merchant, transactionNature) {
  if (!isRewardEligibleNature(transactionNature)) {
    return 0;
  }

  const merchantKey = normalizeCardKey(merchant);
  const categoryKey = normalizeCardKey(category);
  const advancedRules = getRewardRulesForCard(card);
  advancedRules.sort((a, b) => Number(a.priority || 999) - Number(b.priority || 999));

  for (const rule of advancedRules) {
    const ruleScope = rule.scope || inferRewardRuleScope(rule);
    const normalizedRuleCategory = normalizeCardKey(rule.category);
    const categoryMatch = ruleScope === "merchant"
      ? true
      : (!normalizedRuleCategory
        || normalizedRuleCategory === categoryKey
        || normalizedRuleCategory === "all"
        || normalizedRuleCategory === "any"
        || normalizedRuleCategory === "other spends"
        || normalizedRuleCategory === "others"
        || categoryKey.includes(normalizedRuleCategory)
        || normalizedRuleCategory.includes(categoryKey));
    const merchantRule = normalizeCardKey(rule.merchant);
    const merchantMatch = ruleScope === "category"
      ? true
      : (!merchantRule
        || merchantRule === "any"
        || merchantKey.includes(merchantRule)
        || merchantRule.includes(merchantKey)
        || merchantKey.replace(/[()]/g, "").includes(merchantRule.replace(/[()]/g, "")));
    if (!categoryMatch || !merchantMatch) {
      continue;
    }

    const baseUnit = Number(rule.baseUnit || 0);
    const rewardPerUnit = Number(rule.rewardPerUnit || 0);
    const multiplier = Number(rule.multiplier || 1);
    const effectiveRewardPerUnit = rewardPerUnit * multiplier;

    if (baseUnit > 0) {
      return Math.floor(amount / baseUnit) * effectiveRewardPerUnit;
    }
    return amount * effectiveRewardPerUnit;
  }

  const rewardRate = Number(card.rewardRules[category] || card.rewardRules[category.toLowerCase()] || card.rewardRules.default || 0);
  if (!rewardRate) {
    return 0;
  }

  return (amount * rewardRate) / 100;
}

function normalizeAdvancedRewardRule(rule) {
  return {
    id: rule.id || createId(),
    cardName: rule.cardName || rule["Card Name"] || rule.CardName || "",
    category: rule.category || rule.Category || "",
    scope: rule.scope || rule.Scope || "",
    merchant: rule.merchant || rule.Merchant || "",
    baseUnit: Number(rule.baseUnit || rule.BaseUnit || 0),
    rewardPerUnit: Number(rule.rewardPerUnit || rule.RewardPerUnit || 0),
    multiplier: Number(rule.multiplier || rule.Multiplier || 1),
    priority: Number(rule.priority || rule.Priority || 999)
  };
}

function normalizeMilestoneRule(rule) {
  return {
    id: rule.id || createId(),
    cardName: rule.cardName || rule["Card Name"] || rule.CardName || "",
    milestoneType: rule.milestoneType || rule["Milestone Type"] || rule.MilestoneType || rule.type || "",
    cycleType: rule.cycleType || rule["Cycle Type"] || rule.CycleType || "Annual",
    target: Number(rule.target || rule["Annual Spend Requirement"] || rule.annualSpendRequirement || 0),
    rewardType: rule.rewardType || rule["Reward Type"] || rule.RewardType || "",
    rewardValue: Number(rule.rewardValue || rule["Reward Value"] || rule.RewardValue || 0),
    notes: rule.notes || rule.Notes || "",
    reward: rule.reward || rule.notes || rule["Notes"] || "",
    benefit: rule.benefit || rule.notes || rule["Notes"] || ""
  };
}

function normalizeSharedLimitGroup(group) {
  return {
    id: String(group.id || createId()),
    groupName: String(group.groupName || group["Group Name"] || "").trim(),
    totalLimit: Number(group.totalLimit || group["Total Limit"] || 0),
    memberCardIds: Array.isArray(group.memberCardIds)
      ? group.memberCardIds.map(String)
      : splitCsv(group.memberCardIds || group["Member Card Ids"]).map(String),
    notes: String(group.notes || group["Notes"] || "").trim()
  };
}

function inferRewardRuleScope(rule) {
  const merchant = normalizeCardKey(rule.merchant);
  const category = normalizeCardKey(rule.category);
  if ((merchant && merchant !== "any") && (category === "all" || category === "any" || !category)) {
    return "merchant";
  }
  if (!merchant || merchant === "any") {
    return "category";
  }
  return "category-merchant";
}

function resolveRewardRuleMerchantValue(scope, merchantValue) {
  if (scope === "category") {
    return "Any";
  }
  return merchantValue || "Any";
}

function updateRewardRuleScopeUI() {
  if (!elements.rewardRuleScope || !elements.rewardRuleMerchant) {
    return;
  }

  const scope = elements.rewardRuleScope.value;
  if (scope === "category") {
    elements.rewardRuleMerchant.value = "";
    elements.rewardRuleMerchant.disabled = true;
    elements.rewardRuleMerchant.placeholder = "Any merchant in this category";
    return;
  }

  elements.rewardRuleMerchant.disabled = false;
  if (scope === "merchant") {
    elements.rewardRuleCategory.value = elements.rewardRuleCategory.value || "All";
    elements.rewardRuleMerchant.placeholder = "Specific merchant, category optional";
    return;
  }

  elements.rewardRuleMerchant.placeholder = "Specific merchant inside this category";
}

function parseRewardRulesInput(value) {
  const rawInput = String(value || "").trim();
  if (!rawInput) {
    return { rewardRules: {}, rawInput: "" };
  }

  if (rawInput.startsWith("{")) {
    const parsed = JSON.parse(rawInput);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      throw new Error("Reward rules JSON must be an object.");
    }
    return { rewardRules: parsed, rawInput };
  }

  const rewardRules = {};
  rawInput.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).forEach((line) => {
    const percentMatch = line.match(/^(.+?)(?:\:|=)\s*([\d.]+)\s*%?$/i);
    if (percentMatch) {
      rewardRules[percentMatch[1].trim()] = Number(percentMatch[2]);
      return;
    }

    const sentenceMatch = line.match(/^(.+?)\s+([\d.]+)\s*%$/i);
    if (sentenceMatch) {
      rewardRules[sentenceMatch[1].trim()] = Number(sentenceMatch[2]);
    }
  });

  return { rewardRules, rawInput };
}

function parseAdvancedRewardRulesInput(value) {
  const rawInput = String(value || "").trim();
  if (!rawInput) {
    return [];
  }

  if (rawInput.startsWith("[")) {
    const parsed = JSON.parse(rawInput);
    if (!Array.isArray(parsed)) {
      throw new Error("Advanced reward rules must be a JSON array.");
    }
    return parsed;
  }

  throw new Error("Please enter valid JSON for advanced reward rules.");
}

function parseMilestonesInput(value) {
  return parseTargetListInput(value, "reward");
}

function parseAnnualBenefitsInput(value) {
  return parseTargetListInput(value, "benefit");
}

function parseTargetListInput(value, fieldName) {
  const rawInput = String(value || "").trim();
  if (!rawInput) {
    return [];
  }

  if (rawInput.startsWith("[")) {
    const parsed = JSON.parse(rawInput);
    if (!Array.isArray(parsed)) {
      throw new Error("Milestones and annual benefits must be a JSON array.");
    }
    return parsed;
  }

  return rawInput.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).map((line) => {
    const targetMatch = line.match(/(\d[\d,]*)/);
    return {
      target: targetMatch ? Number(targetMatch[1].replace(/,/g, "")) : 0,
      [fieldName]: line
    };
  });
}
function normalizeCard(card) {
  return {
    id: String(card.id || createId()),
    cardName: String(card.cardName || "").trim(),
    bankName: String(card.bankName || inferBankName(card.cardName || "")).trim(),
    creditLimit: Number(card.creditLimit || 0),
    billingCycleDate: Number(card.billingCycleDate || card.statementDate || 1),
    dueDate: Number(card.dueDate || 1),
    statementDate: Number(card.statementDate || card.billingCycleDate || 1),
    anniversaryDate: card.anniversaryDate || new Date().toISOString().split("T")[0],
    renewalTarget: Number(card.renewalTarget || 0),
    sharedGroup: String(card.sharedGroup || card.bankName || "").trim(),
    preferredCategories: Array.isArray(card.preferredCategories) ? card.preferredCategories : splitCsv(card.preferredCategories),
    preferredTransactionTypes: Array.isArray(card.preferredTransactionTypes) ? card.preferredTransactionTypes : splitCsv(card.preferredTransactionTypes),
    supportedTransactionNature: Array.isArray(card.supportedTransactionNature) ? card.supportedTransactionNature : splitCsv(card.supportedTransactionNature),
    rewardRules: typeof card.rewardRules === "object" && card.rewardRules && !Array.isArray(card.rewardRules) ? card.rewardRules : {},
    rewardRulesInput: String(card.rewardRulesInput || "").trim(),
    advancedRewardRules: Array.isArray(card.advancedRewardRules) ? card.advancedRewardRules : [],
    advancedRewardRulesInput: String(card.advancedRewardRulesInput || "").trim(),
    milestones: Array.isArray(card.milestones) ? card.milestones : [],
    milestonesInput: String(card.milestonesInput || "").trim(),
    annualBenefits: Array.isArray(card.annualBenefits) ? card.annualBenefits : [],
    annualBenefitsInput: String(card.annualBenefitsInput || "").trim()
  };
}

function normalizeTransaction(transaction) {
  return {
    id: String(transaction.id || createId()),
    date: transaction.date || new Date().toISOString().split("T")[0],
    cardId: String(transaction.cardId || ""),
    cardName: String(transaction.cardName || "").trim(),
    amount: Number(transaction.amount || 0),
    category: String(transaction.category || "Others"),
    description: String(transaction.description || "").trim(),
    merchant: String(transaction.merchant || "").trim(),
    transactionNature: normalizeTransactionNature(transaction.transactionNature),
    transactionType: String(transaction.transactionType || "Other"),
    paymentStatus: normalizePaymentStatus(transaction.paymentStatus),
    billPdfLink: String(transaction.billPdfLink || "").trim(),
    unbilledAmount: Number(transaction.unbilledAmount || 0),
    notes: String(transaction.notes || "").trim(),
    rewardPoints: Number(transaction.rewardPoints || 0)
  };
}

function normalizeRewardActivity(item) {
  return {
    id: String(item.id || createId()),
    cardId: String(item.cardId || ""),
    cardName: String(item.cardName || "").trim(),
    action: String(item.action || "").trim(),
    points: Number(item.points || 0),
    date: item.date || new Date().toISOString().split("T")[0]
  };
}

function normalizeTransactionNature(value) {
  const normalized = normalizeCardKey(value);
  if (normalized === "payment" || normalized === "credit") return "Payment";
  if (normalized === "refund") return "Refund";
  if (normalized === "bonus") return "Bonus";
  return "Expense";
}

function normalizePaymentStatus(value) {
  const normalized = normalizeCardKey(value);
  if (normalized === "paid") return "Paid";
  if (normalized === "partially paid" || normalized === "partial" || normalized === "partiallypaid") return "Partially Paid";
  return "Unpaid";
}

function isRewardEligibleNature(value) {
  return normalizeTransactionNature(value) === "Expense";
}

function isExpenseNature(value) {
  return normalizeTransactionNature(value) === "Expense";
}

function isCreditNature(value) {
  const normalized = normalizeTransactionNature(value);
  return normalized === "Payment" || normalized === "Refund" || normalized === "Bonus";
}

function getTransactionTone(value) {
  if (isExpenseNature(value)) return "debit";
  if (isCreditNature(value)) return "credit";
  return "neutral";
}

function getNatureLabel(value) {
  const normalized = normalizeTransactionNature(value);
  if (normalized === "Payment") return "Credit / Payment";
  if (normalized === "Refund") return "Credit / Refund";
  if (normalized === "Bonus") return "Credit / Bonus";
  return "Debit / Expense";
}

function getAmountClass(transaction) {
  const tone = getTransactionTone(transaction.transactionNature);
  if (tone === "debit") return "amount-debit";
  if (tone === "credit") return "amount-credit";
  return "amount-neutral";
}

function transactionMatchesCard(transaction, card) {
  if (!transaction || !card) return false;
  if (transaction.cardId && card.id && String(transaction.cardId) === String(card.id)) return true;
  return normalizeCardKey(transaction.cardName) === normalizeCardKey(card.cardName);
}

function rewardActivityMatchesCard(item, card) {
  if (!item || !card) return false;
  if (item.cardId && card.id && String(item.cardId) === String(card.id)) return true;
  return normalizeCardKey(item.cardName) === normalizeCardKey(card.cardName);
}

function getCardById(cardId) {
  return state.cards.find((card) => String(card.id) === String(cardId)) || null;
}

function getCardNameById(cardId) {
  const card = getCardById(cardId);
  return card ? card.cardName : "";
}

function inferBankName(cardName) {
  const value = normalizeCardKey(cardName);
  if (value.startsWith("hdfc")) return "HDFC Bank";
  if (value.startsWith("axis")) return "Axis Bank";
  if (value.startsWith("amex")) return "American Express";
  if (value.startsWith("idfc")) return "IDFC First";
  if (value.startsWith("icici")) return "ICICI Bank";
  if (value.startsWith("kotak")) return "Kotak Bank";
  if (value.startsWith("stancy") || value.startsWith("standard")) return "Standard Chartered";
  if (value.startsWith("rbl")) return "RBL Bank";
  if (value.startsWith("indusind")) return "Indusind Bank";
  if (value.startsWith("sbi")) return "SBI Bank";
  return "Card Account";
}

function getDefaultCategoriesForCard(cardName) {
  const value = normalizeCardKey(cardName);
  if (value.includes("marriott")) return ["Travel", "Food & Dining"];
  if (value.includes("swiggy")) return ["Food & Dining", "Groceries"];
  if (value.includes("amazon")) return ["Shopping", "Bills & Utilities"];
  if (value.includes("bpcl")) return ["Fuel"];
  if (value.includes("indigo")) return ["Travel"];
  return ["Travel", "Shopping"];
}

function getDefaultRewardRulesForCard(cardName) {
  const value = normalizeCardKey(cardName);
  if (value.includes("swiggy")) return { "Food & Dining": 10, Groceries: 5, default: 1 };
  if (value.includes("amazon")) return { Shopping: 5, default: 1 };
  if (value.includes("bpcl")) return { Fuel: 4, default: 1 };
  return { Travel: 2, Shopping: 2, default: 1 };
}

function getDefaultAdvancedRulesForCard(cardName) {
  const value = normalizeCardKey(cardName);
  if (value.includes("marriott")) {
    return [
      { category: "Travel", merchant: "Marriott", baseUnit: 150, rewardPerUnit: 8, multiplier: 1, priority: 1 },
      { category: "Food & Dining", merchant: "", baseUnit: 150, rewardPerUnit: 4, multiplier: 1, priority: 2 }
    ];
  }
  if (value.includes("atlus") || value.includes("atlas")) {
    return [
      { category: "Travel", merchant: "", baseUnit: 100, rewardPerUnit: 2, multiplier: 5, priority: 1 },
      { category: "Food & Dining", merchant: "", baseUnit: 100, rewardPerUnit: 2, multiplier: 1, priority: 2 }
    ];
  }
  return [];
}

function buildOptions(values) {
  return values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
}

function splitCsv(value) {
  return String(value || "").split(",").map((item) => item.trim()).filter(Boolean);
}

function normalizeCardKey(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function createId() {
  return `id-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(Number(value || 0));
}

function formatPoints(value) {
  const amount = Number(value || 0);
  return `${new Intl.NumberFormat("en-IN", {
    minimumFractionDigits: amount % 1 === 0 ? 0 : 2,
    maximumFractionDigits: 2
  }).format(amount)} pts`;
}

function formatDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(date);
}

function formatDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function diffDays(later, earlier) {
  const ms = startOfDay(later).getTime() - startOfDay(earlier).getTime();
  return Math.round(ms / 86400000);
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function createDateForDay(year, monthIndex, dayOfMonth) {
  const safeDay = Math.max(1, Math.min(31, Number(dayOfMonth || 1)));
  const date = new Date(year, monthIndex, 1);
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  date.setDate(Math.min(safeDay, lastDay));
  return date;
}

function formatJsonForEditor(value) {
  return JSON.stringify(value || [], null, 2);
}

function formatRewardRulesForEditor(rewardRules) {
  return Object.entries(rewardRules || {}).map(([key, itemValue]) => `${key}: ${itemValue}%`).join("\n");
}

function getMultiSelectValues(selectElement) {
  if (!selectElement) {
    return [];
  }
  return Array.from(selectElement.selectedOptions || []).map((option) => String(option.value));
}

function setMultiSelectValues(selectElement, values) {
  if (!selectElement) {
    return;
  }
  const selected = new Set((values || []).map(String));
  Array.from(selectElement.options).forEach((option) => {
    option.selected = selected.has(String(option.value));
  });
}

function showToast(message) {
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    elements.toast.classList.remove("show");
  }, 2600);
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
