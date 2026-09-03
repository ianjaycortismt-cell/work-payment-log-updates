/* Defaults for a brand-new install. Nothing personal lives here: a new user
   starts with an empty timesheet and answers a few questions on first run. */
window.WPL_SEED = {
  settings: {
    currency: "EUR",
    currencySymbol: "",        /* only used when currency is "custom" */
    rate: 10,
    defaultStart: "08:00",
    defaultFinish: "16:00",
    breakHours: 0.5,
    firstPayday: "",           /* asked during setup */
    cycleDays: 28,
    weekHours: 40,             /* what a "full week" means on the progress bar */
    appColor: "orange",
    takeHome: {
      country: "MT",
      profile: "mt-single",
      social: "mt-prorata",
      otherIncome: 0,
      workMonths: 12,
      customTax: 0,
      customSocial: 0,
      customOther: 0,
      customFixed: 0
    },
    spendingPlan: {
      incomeMode: "takehome",
      customIncome: 0,
      categories: [
        { id: "bills", name: "Bills", color: "#3b82f6", mode: "percent", value: 35 },
        { id: "food", name: "Food", color: "#f59e0b", mode: "percent", value: 20 },
        { id: "savings", name: "Savings", color: "#10b981", mode: "percent", value: 25 },
        { id: "spending", name: "Spending money", color: "#8b5cf6", mode: "percent", value: 20 }
      ]
    },
    logReminderTime: "18:00",
    logReminderDays: [1, 2, 3, 4, 5],
    onboarded: false
  },

  entries: [],

  /* What each kind of leave pays, as a percentage of a normal day. These are
     only starting points — every one of them is editable in Settings, because
     entitlements differ by country, employer and contract. */
  leaveRates: {
    annual: 100,
    public: 100,
    sick: 100,
    injury: 100,
    quarantine: 100,
    bereavement: 100,
    marriage: 100,
    birth: 100,
    maternity: 100,
    adoption: 100,
    parental: 0,
    carers: 0,
    family: 0,
    jury: 100,
    study: 100,
    toil: 100,
    unpaid: 0,
    other: 0,
    off: 0
  },

  currencies: [
    { code: "EUR", symbol: "€" },
    { code: "GBP", symbol: "£" },
    { code: "USD", symbol: "$" },
    { code: "AUD", symbol: "A$" },
    { code: "CAD", symbol: "C$" },
    { code: "CHF", symbol: "CHF " },
    { code: "SEK", symbol: "kr " },
    { code: "NOK", symbol: "kr " },
    { code: "DKK", symbol: "kr " },
    { code: "PLN", symbol: "zł " },
    { code: "CZK", symbol: "Kč " },
    { code: "INR", symbol: "₹" },
    { code: "ZAR", symbol: "R" },
    { code: "NZD", symbol: "NZ$" },
    { code: "JPY", symbol: "¥" },
    { code: "custom", symbol: "" }
  ],

  notes: [
    "Pay periods end on payday and use the cycle length in Settings.",
    "Only days in your timesheet are counted.",
    "Pay is before tax and other deductions.",
    "The unpaid break is taken off each worked day. No start and finish time means zero hours."
  ]
};
