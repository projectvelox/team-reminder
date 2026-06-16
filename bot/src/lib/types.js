// v1.7.43 — central JSDoc typedefs so // @ts-check on the handler files has
// real types to lean on. Pure documentation; no runtime exports.
//
// Add @ts-check at the top of any file that should be type-checked. The
// repo's bot/jsconfig.json + `npm run typecheck` script run this across
// every file in src/.

// @ts-check

/**
 * @typedef {Object} TeamsUserClaims
 * @property {string} oid           Entra object id (stable per-user)
 * @property {string} [tid]         Entra tenant id
 * @property {string|null} [upn]
 * @property {string|null} [name]
 */

/**
 * @typedef {Object} Reminder
 * @property {string} id
 * @property {string} title
 * @property {string|null} time          HH:MM, or null = anytime
 * @property {boolean} done
 * @property {string|null} firedAt       ISO when proactive card last fired
 * @property {string|null} createdDate   YYYY-MM-DD
 * @property {string|null} closedAt      ISO when marked done
 * @property {string[]} tags
 * @property {'normal'|'high'} priority
 * @property {string|null} dueAt         YYYY-MM-DD
 * @property {string|null} description
 * @property {number} rollDays
 * @property {number} [leadMinutes]      per-reminder override; null = settings default
 * @property {string|null} [snoozedUntil] ISO
 * @property {string|null} [client]
 * @property {'none'|'daily'|'weekdays'|'weekly'} [repeat]
 * @property {Array<{id:string,text:string,done:boolean}>} [subtasks]
 */

/**
 * @typedef {Object} LicenseComment
 * @property {string} id
 * @property {string} at        ISO
 * @property {string} byOid
 * @property {string|null} byName
 * @property {string} text
 */

/**
 * @typedef {Object} LicenseEvent
 * @property {string} at
 * @property {string} byOid
 * @property {string|null} byName
 * @property {'created'|'statusChanged'|'ownerChanged'|'expiryChanged'|'renewed'|'abandoned'|'customerMerged'} type
 * @property {string|null} [detail]
 */

/**
 * @typedef {Object} License
 * @property {string} id
 * @property {string} customer
 * @property {string} licenseType
 * @property {number} userCount
 * @property {string|null} expiryDate     YYYY-MM-DD
 * @property {string|null} ownerOid
 * @property {string|null} ownerName
 * @property {string|null} productLine
 * @property {number[]|null} leadDays
 * @property {number[]} lastFiredLeadDays
 * @property {string|null} notes
 * @property {LicenseComment[]} [comments]
 * @property {'active'|'abandoned'} state
 * @property {'notStarted'|'noticeSent'|'awaitingCustomer'|'customerConfirmed'|'renewed'} status
 * @property {string|null} statusChangedAt
 * @property {string|null} statusChangedByOid
 * @property {string|null} statusChangedByName
 * @property {'annual'|'biennial'|'triennial'} renewalCycle
 * @property {string|null} createdAt
 * @property {string|null} createdByOid
 * @property {string|null} createdByName
 * @property {string|null} lastEditedAt
 * @property {string|null} lastEditedByOid
 * @property {string|null} lastEditedByName
 * @property {string|null} lastRenewedAt
 * @property {string|null} lastFollowUpAt
 * @property {number|null} lastEscalatedDays
 * @property {string|null} leadSnoozedUntil
 * @property {string|null} [deletedAt]      v1.7.43 soft-delete
 * @property {string|null} [deletedByOid]
 * @property {string|null} [deletedByName]
 * @property {LicenseEvent[]} events
 */

/**
 * @typedef {Object} Settings
 * @property {string} eodTime                       HH:MM
 * @property {number} leadMinutes
 * @property {boolean} weekdaysOnly
 * @property {boolean} notifications
 * @property {string|null} quietStart
 * @property {string|null} quietEnd
 * @property {boolean} autoImportFlagged
 * @property {number[]} licenseLeadDays
 * @property {boolean} licenseSkipBriefing
 * @property {boolean} licenseSkipMonthlyDigest
 * @property {boolean} licenseRollupDigest
 * @property {Array<{id:string,name:string,filters:Object}>} savedLicenseViews
 */

/**
 * @typedef {Object} UserRecord
 * @property {Settings} settings
 * @property {Object|null} conversationRef
 * @property {string|null} lastEodDate
 * @property {string|null} lastRolloverDate
 * @property {string|null} tenantId
 * @property {string|null} serviceUrl
 * @property {string|null} displayName
 * @property {string|null} lastBriefingDate
 * @property {string|null} briefingSnoozedUntil
 * @property {string|null} coldNudgedAt
 * @property {string|null} lastDigestSentMonth
 */

module.exports = {};
