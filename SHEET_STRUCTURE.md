# Google Sheet Structure

Create one Google Spreadsheet and add these three sheets.

## 1. `Cards`

| Column | Description |
| --- | --- |
| `id` | Unique card ID from the app |
| `cardName` | Display name of the card |
| `bankName` | Bank name |
| `creditLimit` | Total credit limit |
| `billingCycleDate` | Billing day of month |
| `dueDate` | Payment due day of month |
| `statementDate` | Statement day of month |
| `anniversaryDate` | Card anniversary/start date |
| `renewalTarget` | Annual renewal / spend target |
| `sharedGroup` | Bank or shared card group |
| `preferredCategoriesJson` | JSON array of important categories |
| `preferredTransactionTypesJson` | JSON array like Online, UPI, Wallet |
| `supportedTransactionNatureJson` | JSON array like Expense, Payment, Refund |
| `rewardRulesJson` | JSON object for category reward rates |
| `advancedRewardRulesJson` | JSON array for category/merchant-based reward rules |
| `milestonesJson` | JSON array for cycle-based milestones |
| `annualBenefitsJson` | JSON array for annual spend benefits |
| `createdAt` | Timestamp added by Apps Script |

Example:

```json
{
  "Fuel": 1,
  "Grocery": 2,
  "Travel": 3,
  "Shopping": 5,
  "Others": 1
}
```

Milestones example:

```json
[
  { "target": 100000, "reward": "2000 bonus points" },
  { "target": 300000, "reward": "Annual fee waiver" }
]
```

Advanced reward rules example:

```json
[
  {
    "category": "Food & Dining",
    "merchant": "Swiggy",
    "baseUnit": 150,
    "rewardPerUnit": 5,
    "multiplier": 2,
    "priority": 1
  }
]
```

Annual benefits example:

```json
[
  { "target": 250000, "benefit": "Fee waiver" },
  { "target": 500000, "benefit": "Premium lounge access" }
]
```

## 2. `Transactions`

| Column | Description |
| --- | --- |
| `id` | Unique transaction ID |
| `date` | Purchase date in `YYYY-MM-DD` |
| `cardId` | Linked card ID |
| `cardName` | Card display name |
| `amount` | Transaction amount |
| `category` | Fuel, Grocery, Travel, Shopping, Others |
| `description` | User-friendly transaction description |
| `merchant` | Merchant name |
| `transactionNature` | Expense, Payment, Refund, Bonus |
| `transactionType` | Online, UPI, POS, Wallet, etc. |
| `paymentStatus` | Unpaid, Paid, Partially Paid |
| `billPdfLink` | Optional bill or statement link |
| `unbilledAmount` | Amount still not billed / not settled |
| `notes` | Optional note |
| `rewardPoints` | Points/cashback earned |
| `createdAt` | Timestamp added by Apps Script |

## 3. `RewardActivities`

| Column | Description |
| --- | --- |
| `id` | Unique reward activity ID |
| `cardId` | Linked card ID |
| `cardName` | Card display name |
| `action` | `redeemed` or `expiry` |
| `points` | Number of points |
| `date` | Activity date in `YYYY-MM-DD` |
| `createdAt` | Timestamp added by Apps Script |
