# Security Specification: Event Registration

## 1. Data Invariants
- Each application must contain valid `name` (1-10 Korean/English letters), `phone` (000-0000-0000 format), `email` (valid email format up to 80 chars), `slot` (one of the 3 allowed sessions), `code` (6 alphanumeric characters), and `agree: true`.
- Extra arbitrary fields ("Ghost fields") outside the allowed blueprint schema are strictly rejected.
- Path parameter `applicationId` must be a valid alphanumeric/hyphen/underscore identifier up to 128 chars.
- Code cannot be mutated upon update.

## 2. The "Dirty Dozen" Payloads (Must be rejected)
1. Ghost field injection: `{ name: "홍길동", phone: "010-1234-5678", email: "hong@test.com", slot: "1회차 · 평일 오전", code: "ABC123", agree: true, isAdmin: true }`
2. Name too long (>10 chars): `{ name: "김수한무거북이와두루미삼천갑자", ... }`
3. Invalid characters in name: `{ name: "John123#", ... }`
4. Invalid phone format: `{ phone: "01012345678", ... }`
5. Invalid email format: `{ email: "not-an-email", ... }`
6. Disallowed session slot: `{ slot: "VIP Session", ... }`
7. Consent false: `{ agree: false, ... }`
8. Code length invalid: `{ code: "TOOLONGCODE", ... }`
9. Memo too long (>100 chars): `{ memo: "A".repeat(101), ... }`
10. Missing required field: `{ name: "홍길동", email: "hong@test.com" }`
11. Malicious document ID: `/applications/../../root`
12. Oversized document ID: `/applications/` + "A".repeat(200)
