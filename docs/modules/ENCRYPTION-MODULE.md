# modEncryption — Password Hashing Module

The system's password security layer. Versioned 1.2.22, last touched 01/10/2014 — making this the **oldest module** in the codebase (older than `modCommon` 28/10/2014 and `modDatabase` 17/12/2014). Authored by Aeric Poon, named "for simple encryption" — and the word "simple" is doing a lot of work in that comment.

This is a small module (four functions, ~80 lines) but it's the foundation of every authentication path in the application: login, password change, admin step-up auth, user creation. Every password hash in the database was produced by `GoldFishEncode`.

## What this module provides

Four functions, two public and two private:

1. **`Encrypt(plaintext, salt)`** — public wrapper that applies salt and hashes
2. **`GenSalt(length)`** — public salt generator
3. **`GoldFishEncode(pw)`** — public hash function (the "simple encryption" itself)
4. **`Dec2Bin(dec)`** — private helper for decimal-to-binary conversion

The module is tiny but every form's authentication path depends on it. `frmUserLogin` calls `GoldFishEncode` directly; `frmAdmin` and `frmUserChangePassword` call it directly too; `frmUserMaintain` and `frmUserChangePassword` call `Encrypt` and `GenSalt` for password updates. Inconsistent — some paths inline the salt-and-hash dance, others go through the wrapper.

## Function-by-function

### `Encrypt(strPlaintext, strSalt) As String`

The high-level wrapper. Three-line implementation:

```vb
Encrypt = GoldFishEncode(strPlaintext & strSalt)
```

Concatenates the password and salt (in that order), passes the whole string through `GoldFishEncode`, returns the result. That's the entire algorithm: **salt is appended, not prepended; concatenated, not interleaved; passed once, not iterated.**

This is the function that callers *should* be using consistently — the asymmetry noted in the form analyses (where `frmUserLogin` and `frmAdmin` inline the same `GoldFishEncode(pw & salt)` operation while `frmUserMaintain` and `frmUserChangePassword` use `Encrypt`) is a code-style inconsistency, not a behavioral one. Both code paths produce identical hashes for identical inputs because they perform the same operation.

If anyone ever extends `Encrypt` to do something different — say, add iterations or a peppered constant — the inline callers will silently break. The integration contract is "verify with `GoldFishEncode(pw & salt)`, store with `Encrypt(pw, salt)` — and these must be equivalent forever."

### `GenSalt(StringLen As Integer) As String`

Generates a random salt string. Three things to flag immediately:

```vb
'Max length = 6
If StringLen > 6 Then StringLen = 6
StringLen = StringLen \ 2
For i = 1 To StringLen
    Randomize
    GenSalt = GenSalt & Hex((Rnd() * 64) Mod 100)
Next
```

**Length is silently capped at 6.** Pass any value larger than 6 and it's truncated. Then it's halved (integer divide by 2). So `GenSalt(4)` produces 2 iterations, `GenSalt(6)` produces 3 iterations, `GenSalt(10)` produces 3 iterations. The "length" parameter doesn't mean what callers think it means.

**Each iteration produces a hex string of 1-2 characters.** `Rnd() * 64` is a value 0-64; `Mod 100` keeps it 0-64 (no-op since the input is already < 100); `Hex()` converts to hex. Hex of values 0-64 produces "0" through "40" — so each iteration adds 1 or 2 hex characters. So:

| Caller passes | Iterations | Actual salt length |
|---|---|---|
| `GenSalt(4)` | 2 | 2-4 chars |
| `GenSalt(6)` | 3 | 3-6 chars |
| `GenSalt(10)` | 3 (capped) | 3-6 chars |

The form analyses noted that `frmUserChangePassword` and `frmUserMaintain` use `GenSalt(4)` while `CreateSampleData` uses `GenSalt(6)`. The actual difference is "2-4 character salt" vs "3-6 character salt" — both well below modern salt lengths (16+ bytes recommended).

**`Randomize` is called inside the loop.** This re-seeds the PRNG on every iteration. Combined with VB6's `Randomize` using the system clock as default seed, two consecutive calls to `Randomize` within the same millisecond use the same seed. That said, `Rnd()` is called immediately after, so within a tight loop the time delta between iterations is microscopic — meaning the salts could exhibit reduced randomness if the PRNG uses second-precision seeding. In practice, modern Windows PRNG is finer-grained than that, but the pattern is suspect. `Randomize` should be called once outside the loop.

**The Salt column is `TEXT(50)` in the schema** — the schema reserves room for much longer salts than this function will ever produce. So upgrading `GenSalt` to produce a longer salt would be a one-line change with no schema impact. The 4-character cap exists only here.

### `GoldFishEncode(pw)` — the actual "encryption"

The core hash function. The algorithm walks each character of the password and XORs it (additively, mod 2) against a rolling state that starts as the binary length of the password. Step by step:

```vb
ibin = Dec2Bin(Len(pw))                   ' Initial state = binary password length
For M = 1 To Len(pw)
    xbin = Dec2Bin(Asc(Mid(pw, M, 1)))    ' Current char as 8-bit binary
    ybin = ""
    For i = 1 To 8                         ' XOR (additive mod 2) bit-by-bit
        ybin = ybin & (CInt(Mid(xbin, i, 1)) + CInt(Mid(ibin, i, 1))) Mod 2
    Next
    ' Convert ybin back to integer
    d = 8 : n = 0
    For i = 1 To 8
        B = Mid(ybin, i, 1) * 2 ^ (d - 1)
        n = n + B
        d = d - 1
    Next
    GoldFishEncode = GoldFishEncode & Hex(n)  ' Append hex digits
    ibin = ybin                                ' Roll state forward
Next
```

What this is, mathematically: a feedback-stream cipher where each output byte is the XOR of the current input byte with the previous output byte (with the very first XOR using the password length as the seed). The output is the concatenation of hex representations of each XOR result.

**Properties of this scheme:**

- **It's reversible without the salt.** Given a `GoldFishEncode` output and the password length, you can decode it. Each output byte is XOR of input byte and previous state — and the state evolution is deterministic from the length. So knowing the output + length recovers the password exactly. This is the single biggest property: **it's not a one-way hash, it's a reversible encoding.**

- **Hex is variable-length per byte.** `Hex(0)` = "0" (1 char), `Hex(15)` = "F" (1 char), `Hex(16)` = "10" (2 chars), `Hex(255)` = "FF" (2 chars). So **the output length depends on the byte values**, not just the input length. A password of "AAAA" (4×0x41) produces a different output length than "ÿÿÿÿ" (4×0xFF). This means the database `UserPassword TEXT(50)` could in theory hold up to 25 input characters, but with the 8/10-char input cap this is never tested.

- **No iteration, no key stretching.** A modern password hash (bcrypt, scrypt, argon2) deliberately costs CPU time per evaluation to slow brute-force. `GoldFishEncode` runs in O(n) where n = password length — it's effectively free to compute. A modern GPU can evaluate billions per second. With a 4-character salt and a 10-character password cap, a brute-force is feasible in seconds.

- **The salt is concatenated to the password before encoding.** So `Encrypt("admin", "ABCD")` is `GoldFishEncode("adminABCD")` — the salt extends the input. This salts the *result* (different salts for the same password give different hashes) but doesn't slow the verification.

- **The first character's XOR is against the password length in binary.** A password length of 5 has binary representation `00000101` (8 bits). The first byte's XOR mixes the first char with `00000101`. Two passwords of identical length and identical first character will produce identical first hash bytes — so the prefix entropy depends only on length and first char.

This is **textbook obfuscation, not cryptography**. The function name "GoldFishEncode" honestly hints at it — "encode" not "hash" or "encrypt." It deters casual snooping (someone with database access can't read passwords directly) and ties the hash to a salt (so identical passwords across users have different hashes), but it doesn't withstand any focused attack:

- A leaked database + a few minutes of analysis recovers all passwords
- The output is reversible with just the password length
- No iteration cost means brute-force is fast
- The 4-character salt provides ~2.8 million combinations (alphanumeric hex), trivial to enumerate
- The 8-10 character password cap dramatically shrinks the keyspace

For an internal hotel-POS application running on a single machine with physical access controlled, this is roughly *adequate* — the threat model is "casual operator nosing through the .mdb file" not "attacker with database dump and rainbow tables." For anything internet-facing, it would be unacceptable. The system's `frmUserLogin` advertising `admin/admin` defaults on the login screen suggests the developer was aware that the security model is permissive.

### `Dec2Bin(dec)` — private helper

Converts an Integer to an 8-bit (or 16-bit, etc.) binary string, padded with leading zeros to a multiple of 8.

```vb
Do Until dec = 0
    Bits = dec Mod 2 & Bits
    dec = dec \ 2
Loop
Do Until Len(Bits) Mod 8 = 0
    Bits = "0" & Bits
Loop
```

Standard divide-by-2 conversion, with leading-zero padding. The padding-to-multiple-of-8 means a length of 5 becomes `00000101`, a length of 200 becomes `11001000`, a length of 256 would become `0000000100000000` (16 bits).

**The padding loop is a subtle correctness anchor.** Without it, `Dec2Bin(5)` would return `"101"` (3 chars) and the inner XOR loop `For i = 1 To 8` would crash on `Mid(xbin, 4, 1)` when index is past the string length. With the padding, every output is at least 8 chars and the loop is safe.

**Edge case: Dec2Bin(0).** The `Do Until dec = 0` loop never executes, leaving `Bits = ""`. Then the padding loop's `Len("") Mod 8 = 0` is True (0 mod 8 = 0), so the loop *also* doesn't execute, and the function returns `""`. This is a bug — `Dec2Bin(0)` should return `"00000000"`. In practice, `GoldFishEncode` calls this with `Len(pw)` (always ≥ 1 since empty passwords are blocked at the form layer) and `Asc(...)` (always ≥ 0 but typically ≥ 32 for printable chars). So an actual zero would only occur if someone passed a password containing the NUL character (`Chr(0)`), which the input fields wouldn't permit. The bug exists but isn't triggerable from normal use.

**The "Mod 8" padding accommodates ASCII codes >127** — `Asc(Chr(255))` is 255, which needs 8 bits, exactly. For Unicode characters above 255, the function would still pad to a multiple of 8 (16, 24, etc.) and the calling code's `For i = 1 To 8` loop would only XOR the first 8 bits, silently dropping high-order bits. Another quiet failure mode for edge cases — passwords containing non-ASCII characters (like accented letters) would have their high bits ignored. The MaxLength=10 input fields make this rare; with IME disabled on password fields, near-impossible.

## Notable points and quirks

- **The module is named `modEncryption`** but contains no encryption — only encoding/hashing. The naming is aspirational. Cryptographers would object to calling `GoldFishEncode` "encryption" because it has no separate key (the salt is part of the input, not a key); they'd object to calling it a "hash" because it's reversible. It's most accurately a *salted reversible encoding scheme*.

- **The "GoldFish" name is unexplained.** Probably a developer joke (goldfish = small memory = simple algorithm?). The author signature in the module header is "Aeric Poon" — same as `modDatabase`. The name isn't a reference to any known cipher.

- **`GenSalt` silently caps at 6.** Callers passing 8, 10, 16 all get the same 3-iteration result. This is a foot-gun — anyone trying to "improve security" by increasing salt length would see no effect without reading the implementation. Worth fixing in any cleanup pass: either honor the requested length, or rename to `GenSalt6` to make the cap explicit.

- **`StringLen \ 2` then `Hex()`** means the actual character output is **roughly equal to the input parameter** for normal hex outputs (each iteration produces 1-2 chars). So `GenSalt(4)` averages ~3 chars output, `GenSalt(6)` averages ~4.5 chars. The math is "iterate length÷2 times, each appending 1-2 hex chars." The relationship between parameter and output length is opaque without reading code.

- **`Randomize` inside the loop is wasteful.** Should be called once before the loop. Within the same millisecond, repeated calls to `Randomize` produce the same seed (VB6's default seed is the system timer in seconds with millisecond fractional precision). The PRNG's internal state advances per `Rnd()` call regardless, but re-seeding every iteration is at best a no-op and at worst defeats sequence variability.

- **Hex is uppercase.** VB6's `Hex()` returns uppercase ("FF" not "ff"). So all stored hashes and salts are uppercase. Combined with case-sensitive string comparison in `If rst!UserPassword = strCheck Then`, the verification works because both sides come through the same `Hex()` call.

- **No constant-time comparison.** `If strPassword = gstrUserPassword Then` in `frmUserLogin` uses VB6's standard string comparison, which short-circuits on the first mismatch. In a high-precision timing attack scenario, an attacker could measure response time to deduce hash characters one at a time. For a local hotel app this is irrelevant; for a network-exposed system it would matter.

- **The error handlers in this module are actually inappropriate for a hash function.** `On Error GoTo CheckErr` followed by `MsgBox Err.Number & " - " & Err.Description` means an unexpected error during hashing would pop a dialog box. For a security-critical primitive, it should fail silently and return a sentinel value (or `Err.Raise` to the caller). A MsgBox during login would interrupt the user with an internal error message. In practice the code paths don't fail — `Mid`, `Asc`, `Dec2Bin` all behave deterministically — but the design choice of MsgBox-on-error is wrong for this layer.

- **The `LogErrorDB` calls are commented out.** Same pattern as `modDatabase`: text-file logging stays, database logging is suppressed. Reasoning likely the same — if a hash fails during login, you can't necessarily count on the database connection being established yet.

- **Thread-safety is irrelevant.** VB6 is single-threaded for ApartmentSTA controls, and the application is single-user single-machine. So the lack of thread-safe `Randomize`/`Rnd` doesn't matter. In a multi-threaded environment this code would have issues.

- **The hash output is unbounded by salt length.** If a user has password "12345" and salt "AB", `GoldFishEncode("12345AB")` runs 7 iterations producing up to 14 hex chars. With max password 10 + salt 6 = 16 input chars, max output is 32 hex chars. Comfortably within the `TEXT(50)` schema column.

- **The module is the smallest in the codebase** (~80 lines, 4 functions) but has the **highest blast radius** if compromised. A bug in `GoldFishEncode` invalidates every stored password. A change to its algorithm requires a database migration or a "force all users to change password on next login" rollout. Worth treating with extra care in any refactor — the seam between "stored hash" and "verification hash" must remain consistent across all callers.

- **Modification dates suggest this was the first module written.** 01/10/2014 predates `modCommon` (28/10/2014) and `modDatabase` (17/12/2014). The developer built the security primitives first, then the SQL builder, then the database connection layer. Reasonable bottom-up sequence.

- **`Option Explicit` is set.** Standard hygiene.

- **A modern replacement** would substitute `GoldFishEncode` with bcrypt or argon2id (available via Win32 BCrypt API or .NET interop), increase salt length to 16+ bytes via `CryptGenRandom`, increase password length cap to 64+, and add iteration count tuning. Total effort: probably 200-400 lines of new code plus a one-time migration script that re-hashes on next login. For an internal hotel app the cost-benefit is marginal; for any future productization it would be a prerequisite.
