AES-256-CBC in VBScript
=======================

This tiny project contains an implementation of AES-256-CBC in VBScript.

[![View Source][Source SVG]][src]
[![MIT License][License SVG]][L]

[Source SVG]: https://img.shields.io/badge/view-source-brightgreen.svg
[License SVG]: https://img.shields.io/badge/license-MIT-blue.svg
[src]: aes.vbs

Note that this project does not implement the cryptographic primitives
from scratch. It is only a convenient wrapper around Microsoft's Common
Object Runtime Library (`mscorlib.dll`).


Contents
--------

* [Features](#features)
* [Why?](#why)
* [Demo Script](#demo-script)
* [Demo Output](#demo-output)
* [Example Ciphertexts](#example-ciphertexts)
* [Crypto Properties](#crypto-properties)
* [OpenSSL CLI Implementation](#openssl-cli-implementation)
* [Questions and Answers](#questions-and-answers)
* [References](#references)


Features
--------

Here are some features of the AES-256-CBC implementation in this
project:

- Works with Base64 encoded key.
- Exposes two simple functions to encrypt and decrypt data.
- Initialization vector (IV) is automatically generated. Caller does not
  need to worry about it.
- The `Encrypt()` function returns the IV and ciphertext concatenated
  together as a single string which can be stored and later fed directly
  to the `Decrypt()` function. Again, no need to worry about maintaining
  the IV yourself.
- Can be used with Classic ASP. Just put the entire [source code][src]
  within the ASP `<%` and `%>` delimiters in a file named `aes.inc` and
  include it in ASP files with an `#include` directive like this:

  ```asp
  <!-- #include file="aes.inc" -->
  ```

Note: This is not a crypto library. You can use it like one if all you
want to do is AES-256-CBC encryption/decryption. But it does not support
anything other than AES-256-CBC. If you want a different key size or
cipher mode, you'll have to dive into the [source code][src] and modify
it as per your needs. If you do so, you might find this documentation
useful: [`RijndaelManaged`][RijndaelManaged].


Why?
----

Why not?

Okay, the real story behind writing this implementation involved a
legacy application with a lot of VBScript code. It used an outdated
cipher that needed to be updated. That's what prompted me to create a
wrapper for AES-256-CBC in VBScript. After writing it, I thought it
would be good to share it on the Internet in case someone else is
looking for something like this.


Demo Script
-----------

On a Windows system, enter the following command to load
[`aes.vbs`][src] and execute the functions in it:

```
cscript aesdemo.wsf
```

The Windows Script File (WSF) named [`aesdemo.wsf`][aesdemo] loads
[`aes.vbs`][src] and executes the functions defined in it.

[aesdemo]: aesdemo.wsf


Demo Output
-----------

The output from the demo script looks like this:


```
demoAESKey: CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg=
encrypted1: gNJryUZ+3QPCt9DP4HQhFQ==:YAMmzeDoPgz0ZpmSnxqb2A==
decrypted1: foo
encrypted2: 2dmVWVT++xbgaDq7ktdUNg==:LEd9iJAHo6bhkpkY/CcrlQ==
decrypted2: foo

aes.BlockSize: 128
aes.FeedbackSize: 128
aes.KeySize: 256
aes.Mode: 1
aes.Padding: 2
aesEnc.InputBlockSize: 16
aesEnc.OutputBlockSize: 16
aesDec.InputBlockSize: 16
aesDec.OutputBlockSize: 16
b64Enc.InputBlockSize: 3
b64Eec.OutputBlockSize: 4
b64Dec.InputBlockSize: 1
b64Dec.OutputBlockSize: 3
```

Only the second line of output (`encrypted1`) changes on every run
because it depends on a dynamically generated initialization vector (IV)
that is used in the encryption. This is, in fact, an important security
requirement for achieving semantic security.

For example, if a database contains encrypted secrets from various users
and if the same plaintext always encrypts to the same ciphertext (which
would happen if the key and IV remain constant between encryptions),
then we can tell that two users have the same plaintext secrets if their
ciphertexts are equal in the database. Being able to do so violates the
requirements of semantic security which requires that an adversary must
not be able to compute any information about a plaintext from its
ciphertext. This is why the ciphertext needs to be different for the
same plaintext in different encryptions and this is why we need a random
IV for each encryption.

There are two blocks of output shown above. The first block shows an
example key and example ciphertexts along with the plaintexts they
decrypt to. The second block shows the default properties of the
cryptography objects used in the VBScript implementation. Both blocks of
output are explained in detail in the sections below.


Example Ciphertexts
-------------------

The first block of output in the demo script looks like this:

```
demoAESKey: CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg=
encrypted1: gNJryUZ+3QPCt9DP4HQhFQ==:YAMmzeDoPgz0ZpmSnxqb2A==
decrypted1: foo
encrypted2: 2dmVWVT++xbgaDq7ktdUNg==:LEd9iJAHo6bhkpkY/CcrlQ==
decrypted2: foo
```

The second line of output (`encrypted1`) changes on every run because it
depends on a dynamically generated initialization vector (IV) used in
the encryption. The fourth line of output (`encrypted2`) remains the
same because it is hardcoded in the demo script.

Each encrypted value in the output is actually a concatenation of the
initialization vector (IV), colon (`:`), and ciphertext.

The ciphertext is generated with a function call like this:

```vbs
demoPlaintext = "foo"
demoAESKey = "CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg="
encrypted1 = Encrypt(demoPlaintext, demoAESKey)
```

The encrypted value (i.e., concatenated IV, colon, and ciphertext) can
be decrypted back to plaintext with a function call like this:

```vbs
decrypted1 = Decrypt(encrypted1, demoAESKey)
```

The `Encrypt` and `Decrypt` functions are defined in [`aes.vbs`][src].


Crypto Properties
-----------------

The second block of output from the demo script looks like this:

```
aes.BlockSize: 128
aes.FeedbackSize: 128
aes.KeySize: 256
aes.Mode: 1
aes.Padding: 2
aesEnc.InputBlockSize: 16
aesEnc.OutputBlockSize: 16
aesDec.InputBlockSize: 16
aesDec.OutputBlockSize: 16
b64Enc.InputBlockSize: 3
b64Eec.OutputBlockSize: 4
b64Dec.InputBlockSize: 1
b64Dec.OutputBlockSize: 3
```

These are the values of the properties of various cryptography objects
used in [`aes.vbs`][src]. The output shows the following details:

  - The `aes` object (of class `RijndaelManaged`) has a default key size
    of 256 bits.
  - The block size is 128 bits. The mode is 1, i.e., CBC. See
    [`RijndaelManaged.Mode`] and [`CipherMode`] documentation to confirm
    that the default mode is indeed CBC.
  - The padding mode is 2, i.e., PKCS #7. See
    [`RijndaelManaged.Padding`] and [`PaddingMode`] documentation to
    confirm that the default padding mode is indeed PKCS #7.

We do not change these defaults in [`aes.vbs`][src] and we supply a
256-bit encryption key to `Encrypt` and `Decrypt` functions to ensure
that we use AES-256-CBC for encryption.

[`RijndaelManaged.Mode`]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rijndaelmanaged.mode
[`CipherMode`]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.ciphermode
[`RijndaelManaged.Padding`]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rijndaelmanaged.padding
[`PaddingMode`]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.paddingmode


OpenSSL CLI Implementation
--------------------------

For troubleshooting purpose, there are two shell scripts named
[`encrypt`](encrypt) and [`decrypt`](decrypt) present in the current
directory. Here is the synopsis of these scripts:

```shell
sh encrypt PLAINTEXT KEY IV
sh decrypt CIPHERTEXT KEY IV
```

These scripts are merely wrappers around OpenSSL. They accept Base64 key
and Base64 IV and convert them to hexadecimal for OpenSSL command line
tool to consume.

To see example usage of the script we will use these three values from
the demo output:

```
demoAESKey: CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg=
encrypted2: 2dmVWVT++xbgaDq7ktdUNg==:LEd9iJAHo6bhkpkY/CcrlQ==
decrypted2: foo
```

Here is how to use the [`encrypt`](encrypt) and [`decrypt`](decrypt)
scripts:

 1. Encrypt the plaintext `foo` with the demo key and IV displayed in
    the output above.

    ```shell
    sh encrypt foo CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg= 2dmVWVT++xbgaDq7ktdUNg==
    ```

    The output should match the ciphertext in the `encrypted2` value.

 2. Decrypt the plaintext from the previous command using the same demo
    AES key and IV.

    ```shell
    sh decrypt LEd9iJAHo6bhkpkY/CcrlQ== CKkPfmeHzhuGf2WYY2CIo5C6aGCyM5JR8gTaaI0IRJg= 2dmVWVT++xbgaDq7ktdUNg==
    ```

    The output should match the plaintext in the `decrypted2` value.


Questions and Answers
---------------------

 1. Can we not use [`AesManaged`][AesManaged] instead of
    [`RijndaelManaged`][RijndaelManaged] in [`aes.vbs`][src] to
    implement AES cipher in VBScript?

    No, we cannot use [`AesManaged`][AesManaged] in VBScript. We can use
    only those classes that are defined in `mscorlib.dll`.
    [`RijndaelManaged`][RijndaelManaged] exists in `mscorlib.dll` but
    [`AesManaged`][AesManaged] does not.

    As a result, we get an error for this VBScript code:

    ```vbs
    set aes = CreateObject("System.Security.Cryptography.AesManaged")
    ```

    Here is the error that occurs:

    ```
    ActiveX component can't create object: 'System.Security.Cryptography.AesManaged'
    ```

    Here is a command that shows the DLLs in which the two classes are
    present:

    ```cmd
    C:\>powershell -Command [System.Reflection.Assembly]::GetAssembly('System.Security.Cryptography.RijndaelManaged')

    GAC    Version        Location
    ---    -------        --------
    True   v4.0.30319     C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll

    C:\>powershell -Command [System.Reflection.Assembly]::GetAssembly('System.Security.Cryptography.AesManaged')

    GAC    Version        Location
    ---    -------        --------
    True   v4.0.30319     C:\windows\Microsoft.Net\assembly\GAC_MSIL\System.Core\v4.0_4.0.0.0__b77a5c561934e089\System.Core.dll
    ```

    This is why we use [`RijndaelManaged`][RijndaelManaged] in our
    VBScript code.

 2. Is [`RijndaelManaged`][RijndaelManaged] suitable for AES cipher
    implementation?

    Yes, it is.

    The [`RijndaelManaged`][RijndaelManaged] class allows setting its
    `KeySize` and `BlockSize` properties independently to `128`, `160`,
    `192`, `224`, or `256`. Further, the key size need not match the
    block size.

    But the [`AesManaged`][AesManaged] class allows setting its
    `KeySize` to `128`, `192`, or `256` only. The `BlockSize` property
    must be `128` only. These constraints match the key sizes and the
    block size defined in the AES standard (see section 5 of [FIPS 197]).

    Therefore, [`AesManaged`][AesManaged] class is the preferred class
    to use because it is not possible to use it in a non-conformant
    manner. But since this class cannot be used in VBScript as explained
    in the previous point and since we stick to the default key size of
    256 bits and the default block size of 128 bits while using
    [`RijndaelManaged`][RijndaelManaged] as explained in section
    [Crypto Properties](#crypto-properties), it is a suitable AES cipher
    implementation.

    Further, the section
    [OpenSSL CLI Implementation](#openssl-cli-implementation) shows that
    the default properties of this cipher is compatible with
    `openssl aes-256-cbc`.

 3. Instead of using a Base64 encoded 256-bit key, can we not use a
    password and derive a 256-bit key from it using a key derivative
    function (KDF)?

    This is not possible because there are only two implementations of
    key derivative functions:

      - [`PasswordDeriveBytes`][PasswordDeriveBytes]: This is an
        implementation of an extension of the PBKDF1 algorithm.
      - [`Rfc2898DeriveBytes`][Rfc2898DeriveBytes]: This is an
        implementation of the PBKDF2 algorithm.

    Although both of these are available in `mscorlib.dll`, they are not
    registered in Windows registry by default. Therefore, like
    [`AesManaged`] class, these two classes too cannot be used in
    VBScript code without modifying the registry.

    If we consider brute-force search of the key, then in case of a key
    derived from a password using a KDF, an attacker could either search
    the password which would be inefficient due to the use of KDF or an
    attacker could directly search the encryption key directly. The
    former (brute-force search of password) is worse than our case where
    we use a 256-bit key directly because with a small password and a
    KDF with small number of iterations, a brute-force search of
    password would require less computational time than brute-force
    search of a 256-bit key. The latter (brute-force search of key) is
    equivalent to our case where we use a Base64 encoded 256-bit key
    because in both cases an attacker has to search for a 256-bit key.
    Therefore, using a 256-bit key directly is equal or better than
    using a key derived from a password.


References
----------

- [Windows Scripting][winscript]
- [VBScript User's Guide][vbsguide]
- [Using Windows Script Files (.wsf)][wsfusage]
- [`cscript`][cscript]
- [`System.Text.UTF8Encoding`][UTF8Encoding]
- [`System.Security.Cryptography.FromBase64Transform`][FromBase64Transform]
- [`System.Security.Cryptography.ToBase64Transform`][ToBase64Transform]
- [`System.Security.Cryptography.RijndaelManaged`][RijndaelManaged]
- [FIPS 197]

[winscript]: https://docs.microsoft.com/en-us/previous-versions/ms950396(v=msdn.10)
[vbsguide]: https://docs.microsoft.com/en-us/previous-versions//sx7b3k7y(v=vs.85)
[wsfusage]: https://docs.microsoft.com/en-us/previous-versions//15x4407c%28v%3dvs.85%29
[cscript]: https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/cscript
[UTF8Encoding]: https://docs.microsoft.com/en-us/dotnet/api/system.text.utf8encoding
[FromBase64Transform]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.frombase64transform
[ToBase64Transform]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.tobase64transform
[RijndaelManaged]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rijndaelmanaged
[AesManaged]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.aesmanaged
[PasswordDeriveBytes]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.passwordderivebytes
[Rfc2898DeriveBytes]: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rfc2898derivebytes
[FIPS 197]: https://nvlpubs.nist.gov/nistpubs/FIPS/NIST.FIPS.197.pdf


License
-------

This is free and open source software. You can use, copy, modify,
merge, publish, distribute, sublicense, and/or sell copies of it,
under the terms of the MIT License. See [LICENSE.md][L] for details.

This software is provided "AS IS", WITHOUT WARRANTY OF ANY KIND,
express or implied. See [LICENSE.md][L] for details.

[L]: LICENSE.md


Support
-------

To report bugs, suggest improvements, or ask questions, please create a
new issue at <http://github.com/susam/aesvbs/issues>.
