#!/usr/bin/env python3
"""A second implementation of Office file encryption, for the tests only.

The engine encrypts and decrypts with its own code; opening its own files with
that same code proves little. This script does the same work with other tools:
Python's hashlib and hmac for the digests, the openssl command line for AES,
RC4 written out here, and its own reader and writer of compound files. The
tests give it files the engine wrote and read back what it writes.

It was written from the same reading of MS-OFFCRYPTO, MS-XLS and MS-DOC as the
engine: it catches mistakes in the arithmetic and in the containers, not a
rule that both got wrong.

    offcrypto_ref.py <command> [options] < input > output

Passwords are given as the hexadecimal form of their UTF-16LE bytes. Exit
codes: 0 done, 2 wrong password, 3 the integrity check failed, 4 the payload
is not a sound ZIP file, 5 anything else.
"""

import hashlib
import hmac
import io
import json
import struct
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from base64 import b64decode, b64encode

WRONG_PASSWORD, BAD_HMAC, BAD_ZIP, OTHER = 2, 3, 4, 5


class Failure(Exception):
    def __init__(self, code, text):
        super().__init__(text)
        self.code = code


# ---------------------------------------------------------------- AES: openssl

def aes(mode, key, data, iv=None, decrypt=False):
    if not data:
        return b''
    args = ['openssl', 'enc', '-aes-%d-%s' % (len(key) * 8, mode), '-K', key.hex(), '-nopad']
    if iv is not None:
        args += ['-iv', iv.hex()]
    if decrypt:
        args.append('-d')
    done = subprocess.run(args, input=data, capture_output=True)
    if done.returncode != 0:
        raise Failure(OTHER, 'openssl failed: ' + done.stderr.decode('utf-8', 'replace'))
    return done.stdout


def rc4(key, data, skip=0):
    box = list(range(256))
    j = 0
    for i in range(256):
        j = (j + box[i] + key[i % len(key)]) & 255
        box[i], box[j] = box[j], box[i]
    i = j = 0
    out = bytearray()
    for n in range(skip + len(data)):
        i = (i + 1) & 255
        j = (j + box[i]) & 255
        box[i], box[j] = box[j], box[i]
        if n >= skip:
            out.append(data[n - skip] ^ box[(box[i] + box[j]) & 255])
    return bytes(out)


# ------------------------------------------------------------- compound files

FREE, END, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD


def cfb_read(data, strict=False):
    """{path: bytes} of every stream, and the list of storages. `strict` also checks the order and the colours of the directory trees."""
    if data[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
        raise Failure(OTHER, 'not a compound file')
    major, shift = struct.unpack_from('<HxxH', data, 26)
    size = 1 << shift
    n_fat, dir_start = struct.unpack_from('<II', data, 44)
    cutoff, minifat_start, n_minifat, difat_start, n_difat = struct.unpack_from('<IIIII', data, 56)
    sector = lambda n: data[(n + 1) * size:(n + 2) * size]
    fat_sectors = [s for s in struct.unpack_from('<109I', data, 76) if s < FATSECT]
    seen = 0
    while difat_start < FATSECT and seen <= len(data) // size:
        entries = struct.unpack('<%dI' % (size // 4), sector(difat_start))
        fat_sectors += [s for s in entries[:-1] if s < FATSECT]
        difat_start = entries[-1]
        seen += 1
    if len(fat_sectors) != n_fat:
        raise Failure(OTHER, 'the header counts %d FAT sectors, the DIFAT lists %d' % (n_fat, len(fat_sectors)))
    if seen != n_difat:
        raise Failure(OTHER, 'the header counts %d DIFAT sectors, the chain has %d' % (n_difat, seen))
    fat = []
    for s in fat_sectors:
        fat += struct.unpack('<%dI' % (size // 4), sector(s))
    for s in fat_sectors:
        if fat[s] != FATSECT:
            raise Failure(OTHER, 'FAT sector %d is not marked as one' % s)

    def chain(start, table):
        out, steps = [], 0
        while start < FATSECT:
            out.append(start)
            start = table[start]
            steps += 1
            if steps > len(table):
                raise Failure(OTHER, 'a sector chain loops')
        if start != END:
            raise Failure(OTHER, 'a sector chain does not end with ENDOFCHAIN')
        return out

    read = lambda start: b''.join(sector(s) for s in chain(start, fat))
    directory = read(dir_start)
    minifat_raw = read(minifat_start)
    minifat = list(struct.unpack('<%dI' % (len(minifat_raw) // 4), minifat_raw))
    if len(minifat_raw) // size != n_minifat:
        raise Failure(OTHER, 'the header counts %d mini FAT sectors, the chain has %d' % (n_minifat, len(minifat_raw) // size))
    entries = []
    for p in range(0, len(directory), 128):
        raw = directory[p:p + 128]
        name_size, kind, colour, left, right, child = struct.unpack_from('<HBBIII', raw, 64)
        start, low, high = struct.unpack_from('<III', raw, 116)
        name = raw[:max(0, name_size - 2)].decode('utf-16-le')
        entries.append(dict(name=name, kind=kind, colour=colour, left=left, right=right, child=child, start=start, size=low))
    root = entries[0]
    if root['kind'] != 5:
        raise Failure(OTHER, 'the first directory entry is not the root')
    mini = read(root['start'])[:root['size']]
    streams, storages, order = {}, [], []

    def key(name):
        return (len(name.encode('utf-16-le')), name.upper())

    def walk(index, prefix, low, high, black_depths, depth, parent_red):
        """In-order walk that checks the order of the names and the colours of the tree."""
        if index == FREE:
            black_depths.append(depth)
            return
        e = entries[index]
        k = key(e['name'])
        if strict and ((low is not None and not low < k) or (high is not None and not k < high)):
            raise Failure(OTHER, 'the directory tree is out of order at "%s"' % e['name'])
        red = e['colour'] == 0
        if strict and red and parent_red:
            raise Failure(OTHER, 'a red directory entry has a red child: "%s"' % e['name'])
        walk(e['left'], prefix, low, k, black_depths, depth + (0 if red else 1), red)
        path = prefix + e['name']
        order.append(path)
        if e['kind'] == 1:
            storages.append(path)
            depths = []
            if strict and e['child'] != FREE and entries[e['child']]['colour'] == 0:
                raise Failure(OTHER, 'the root of the tree under "%s" is red' % path)
            walk(e['child'], path + '/', None, None, depths, 0, False)
            if strict and len(set(depths)) > 1:
                raise Failure(OTHER, 'the tree under "%s" is not balanced' % path)
        elif e['kind'] == 2:
            if e['size'] < cutoff:
                body = b''.join(mini[s * 64:(s + 1) * 64] for s in chain(e['start'], minifat))
            else:
                body = read(e['start'])
            if len(body) < e['size']:
                raise Failure(OTHER, 'the stream "%s" is shorter than its size' % path)
            streams[path] = body[:e['size']]
        walk(e['right'], prefix, k, high, black_depths, depth + (0 if red else 1), red)

    depths = []
    if strict and root['child'] != FREE and entries[root['child']]['colour'] == 0:
        raise Failure(OTHER, 'the root of a directory tree is red')
    walk(root['child'], '', None, None, depths, 0, False)
    if strict and len(set(depths)) > 1:
        raise Failure(OTHER, 'the tree under the root is not balanced')
    return streams, storages, dict(version=major, sector=size, cutoff=cutoff, fat_sectors=n_fat, difat_sectors=n_difat, order=order)


def cfb_write(streams):
    """A version 3 compound file of root-level streams ({name: bytes}); small streams go into the mini stream."""
    names = sorted(streams, key=lambda n: (len(n.encode('utf-16-le')), n.upper()))
    small = [n for n in names if len(streams[n]) < 4096]
    big = [n for n in names if len(streams[n]) >= 4096]
    mini = bytearray()
    mini_start = {}
    minifat = []
    for n in small:
        count = (len(streams[n]) + 63) // 64
        mini_start[n] = len(minifat) if count else END
        minifat += [len(minifat) + i + 1 for i in range(count - 1)] + ([END] if count else [])
        mini += streams[n] + bytes(count * 64 - len(streams[n]))
    sectors_of = lambda size: (size + 511) // 512
    n_dir = sectors_of((len(names) + 1) * 128)
    n_minifat = sectors_of(len(minifat) * 4)
    n_mini = sectors_of(len(mini))
    body = n_dir + n_minifat + n_mini + sum(sectors_of(len(streams[n])) for n in big)
    n_fat = 1
    while n_fat * 128 < n_fat + body:
        n_fat += 1
    if n_fat > 109:
        raise Failure(OTHER, 'too large for this writer')
    fat = [FATSECT] * n_fat
    at = n_fat

    def place(count):
        nonlocal at
        start = at if count else END
        for i in range(count):
            fat.append(at + 1 if i < count - 1 else END)
            at += 1
        return start

    dir_start = place(n_dir)
    minifat_start = place(n_minifat)
    mini_sector = place(n_mini)
    big_start = {n: place(sectors_of(len(streams[n]))) for n in big}
    fat += [FREE] * (n_fat * 128 - len(fat))

    header = bytearray(512)
    header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH', header, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into('<III', header, 40, 0, n_fat, dir_start)
    struct.pack_into('<IIIII', header, 56, 4096, minifat_start, n_minifat, END, 0)
    struct.pack_into('<109I', header, 76, *(list(range(n_fat)) + [FREE] * (109 - n_fat)))

    def entry(name, kind, child, right, start, size):
        raw = bytearray(128)
        text = name.encode('utf-16-le')
        raw[:len(text)] = text
        struct.pack_into('<HBBIII', raw, 64, len(text) + 2, kind, 1, FREE, right, child)
        struct.pack_into('<II', raw, 116, start, size)
        return bytes(raw)

    # The siblings as a list: every entry is black and has only a right child. Not a balanced tree, and readers do not mind.
    directory = entry('Root Entry', 5, 1 if names else FREE, FREE, mini_sector, len(mini))
    for i, n in enumerate(names):
        start = mini_start[n] if n in mini_start else big_start[n]
        directory += entry(n, 2, FREE, i + 2 if i + 1 < len(names) else FREE, start, len(streams[n]))
    while len(directory) % 512:
        blank = bytearray(128)
        struct.pack_into('<III', blank, 68, FREE, FREE, FREE)
        directory += bytes(blank)
    pad = lambda b: bytes(b) + bytes(-len(b) % 512)
    out = bytes(header) + struct.pack('<%dI' % len(fat), *fat) + directory
    out += pad(struct.pack('<%dI' % len(minifat), *minifat) + b'\xff' * (-len(minifat) * 4 % 512))
    out += pad(mini)
    for n in big:
        out += pad(streams[n])
    return out


# ------------------------------------------------------- ECMA-376: agile

NS = '{http://schemas.microsoft.com/office/2006/encryption}'
NSP = '{http://schemas.microsoft.com/office/2006/keyEncryptor/password}'
BLOCK_VERIFIER_INPUT = bytes.fromhex('fea7d2763b4b9e79')
BLOCK_VERIFIER_VALUE = bytes.fromhex('d7aa0f6d3061344e')
BLOCK_KEY_VALUE = bytes.fromhex('146e0be7abacd0d6')
BLOCK_HMAC_KEY = bytes.fromhex('5fb2ad010cb9e1f6')
BLOCK_HMAC_VALUE = bytes.fromhex('a0677f02b22c8433')
PY_HASH = {'SHA1': 'sha1', 'SHA256': 'sha256', 'SHA384': 'sha384', 'SHA512': 'sha512', 'MD5': 'md5'}


def fit(data, size, fill=0x36):
    return data[:size] if len(data) >= size else data + bytes([fill]) * (size - len(data))


def spin(hash_name, salt, password, count):
    h = hashlib.new(hash_name, salt + password).digest()
    for i in range(count):
        h = hashlib.new(hash_name, struct.pack('<I', i) + h).digest()
    return h


def agile_decrypt(streams, password):
    info = streams['EncryptionInfo']
    root = ET.fromstring(info[8:].decode('utf-8'))
    key_data = root.find(NS + 'keyData').attrib
    integrity = root.find(NS + 'dataIntegrity')
    enc = root.find('.//' + NSP + 'encryptedKey').attrib
    if enc['cipherAlgorithm'] != 'AES' or enc['cipherChaining'] != 'ChainingModeCBC':
        raise Failure(OTHER, 'only AES in CBC mode')
    hash_name = PY_HASH[enc['hashAlgorithm']]
    salt = b64decode(enc['saltValue'])
    block = int(enc['blockSize'])
    base = spin(hash_name, salt, password, int(enc['spinCount']))
    key_for = lambda block_key: fit(hashlib.new(hash_name, base + block_key).digest(), int(enc['keyBits']) // 8)
    iv = fit(salt, block)
    verifier_input = aes('cbc', key_for(BLOCK_VERIFIER_INPUT), b64decode(enc['encryptedVerifierHashInput']), iv, True)[:int(enc['saltSize'])]
    verifier_value = aes('cbc', key_for(BLOCK_VERIFIER_VALUE), b64decode(enc['encryptedVerifierHashValue']), iv, True)[:int(enc['hashSize'])]
    if not hmac.compare_digest(hashlib.new(hash_name, verifier_input).digest(), verifier_value):
        raise Failure(WRONG_PASSWORD, 'wrong password')
    secret = aes('cbc', key_for(BLOCK_KEY_VALUE), b64decode(enc['encryptedKeyValue']), iv, True)[:int(key_data['keyBits']) // 8]

    data_hash = PY_HASH[key_data['hashAlgorithm']]
    data_salt = b64decode(key_data['saltValue'])
    data_block = int(key_data['blockSize'])
    package = streams['EncryptedPackage']
    size = struct.unpack_from('<Q', package, 0)[0]
    plain = b''
    for n, at in enumerate(range(8, len(package), 4096)):
        seg_iv = fit(hashlib.new(data_hash, data_salt + struct.pack('<I', n)).digest(), data_block)
        plain += aes('cbc', secret, package[at:at + 4096], seg_iv, True)
    if size > len(plain):
        raise Failure(OTHER, 'the package is shorter than its size')
    plain = plain[:size]

    hmac_ok = None
    if integrity is not None:
        hash_size = int(key_data['hashSize'])
        iv_key = fit(hashlib.new(data_hash, data_salt + BLOCK_HMAC_KEY).digest(), data_block)
        iv_value = fit(hashlib.new(data_hash, data_salt + BLOCK_HMAC_VALUE).digest(), data_block)
        hmac_key = aes('cbc', secret, b64decode(integrity.attrib['encryptedHmacKey']), iv_key, True)[:hash_size]
        hmac_value = aes('cbc', secret, b64decode(integrity.attrib['encryptedHmacValue']), iv_value, True)[:hash_size]
        hmac_ok = hmac.compare_digest(hmac.new(hmac_key, package, data_hash).digest(), hmac_value)
    return plain, hmac_ok


def agile_encrypt(plain, password, hash_label, key_bits, spin_count, seed, password_bits=None):
    """`seed` makes the salts and the key: the output is the same for the same input (the tests keep known answers)."""
    hash_name = PY_HASH[hash_label]
    hash_size = hashlib.new(hash_name).digest_size
    stream = lambda label, size: hashlib.shake_256(seed + label).digest(size)
    salt, data_salt = stream(b'password salt', 16), stream(b'data salt', 16)
    secret, verifier, hmac_key = stream(b'secret key', key_bits // 8), stream(b'verifier', 16), stream(b'hmac key', hash_size)
    # The key of the package may be encrypted with a key of another size than its own.
    password_bits = password_bits or key_bits
    base = spin(hash_name, salt, password, spin_count)
    key_for = lambda block_key: fit(hashlib.new(hash_name, base + block_key).digest(), password_bits // 8)
    block16 = lambda b: b + bytes(-len(b) % 16)
    enc_input = aes('cbc', key_for(BLOCK_VERIFIER_INPUT), block16(verifier), salt)
    enc_value = aes('cbc', key_for(BLOCK_VERIFIER_VALUE), block16(hashlib.new(hash_name, verifier).digest()), salt)
    enc_key = aes('cbc', key_for(BLOCK_KEY_VALUE), block16(secret), salt)

    package = struct.pack('<Q', len(plain))
    for n, at in enumerate(range(0, len(plain), 4096)):
        seg_iv = fit(hashlib.new(hash_name, data_salt + struct.pack('<I', n)).digest(), 16)
        package += aes('cbc', secret, block16(plain[at:at + 4096]), seg_iv)
    iv_key = fit(hashlib.new(hash_name, data_salt + BLOCK_HMAC_KEY).digest(), 16)
    iv_value = fit(hashlib.new(hash_name, data_salt + BLOCK_HMAC_VALUE).digest(), 16)
    enc_hmac_key = aes('cbc', secret, block16(hmac_key), iv_key)
    enc_hmac_value = aes('cbc', secret, block16(hmac.new(hmac_key, package, hash_name).digest()), iv_value)

    b64 = lambda b: b64encode(b).decode('ascii')
    common = 'saltSize="16" blockSize="16" keyBits="%d" hashSize="%d" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="%s"' % (key_bits, hash_size, hash_label)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
           '<keyData %s saltValue="%s"/><dataIntegrity encryptedHmacKey="%s" encryptedHmacValue="%s"/>'
           '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
           '<p:encryptedKey spinCount="%d" %s saltValue="%s" encryptedVerifierHashInput="%s" encryptedVerifierHashValue="%s" encryptedKeyValue="%s"/>'
           '</keyEncryptor></keyEncryptors></encryption>') % (common, b64(data_salt), b64(enc_hmac_key), b64(enc_hmac_value), spin_count, common.replace('keyBits="%d"' % key_bits, 'keyBits="%d"' % password_bits), b64(salt), b64(enc_input), b64(enc_value), b64(enc_key))
    info = struct.pack('<HHI', 4, 4, 0x40) + xml.encode('utf-8')
    return cfb_write({'EncryptionInfo': info, 'EncryptedPackage': package})


# ---------------------------------------------------- ECMA-376: standard

AES_ALG = {128: 0x660E, 192: 0x660F, 256: 0x6610}


def standard_key(salt, password, key_bytes, rounds=50000):
    h = spin('sha1', salt, password, rounds)
    h = hashlib.sha1(h + struct.pack('<I', 0)).digest()
    x1 = hashlib.sha1(bytes(b ^ 0x36 for b in h.ljust(64, b'\0'))).digest()
    x2 = hashlib.sha1(bytes(b ^ 0x5C for b in h.ljust(64, b'\0'))).digest()
    return (x1 + x2)[:key_bytes]


def standard_parse(info):
    major, minor, flags, header_size = struct.unpack_from('<HHII', info, 0)
    header = info[12:12 + header_size]
    h_flags, size_extra, alg, alg_hash, key_bits, provider = struct.unpack_from('<IIIIII', header, 0)
    at = 12 + header_size
    salt_size = struct.unpack_from('<I', info, at)[0]
    salt = info[at + 4:at + 4 + salt_size]
    at += 4 + salt_size
    enc_verifier = info[at:at + 16]
    hash_size = struct.unpack_from('<I', info, at + 16)[0]
    enc_hash = info[at + 20:]
    return dict(major=major, minor=minor, flags=flags, alg=alg, alg_hash=alg_hash, key_bits=key_bits, salt=salt, enc_verifier=enc_verifier, hash_size=hash_size, enc_hash=enc_hash)


def standard_decrypt(streams, password):
    p = standard_parse(streams['EncryptionInfo'])
    if p['alg'] not in AES_ALG.values():
        raise Failure(OTHER, 'only AES')
    key = standard_key(p['salt'], password, p['key_bits'] // 8)
    verifier = aes('ecb', key, p['enc_verifier'], None, True)
    expected = aes('ecb', key, p['enc_hash'][:32], None, True)[:p['hash_size']]
    if not hmac.compare_digest(hashlib.sha1(verifier).digest(), expected):
        raise Failure(WRONG_PASSWORD, 'wrong password')
    package = streams['EncryptedPackage']
    size = struct.unpack_from('<Q', package, 0)[0]
    body = package[8:]
    plain = aes('ecb', key, body[:len(body) // 16 * 16], None, True)
    if size > len(plain):
        raise Failure(OTHER, 'the package is shorter than its size')
    return plain[:size], None


def standard_encrypt(plain, password, key_bits, seed, rounds=50000):
    salt = hashlib.shake_256(seed + b'salt').digest(16)
    verifier = hashlib.shake_256(seed + b'verifier').digest(16)
    key = standard_key(salt, password, key_bits // 8, rounds)
    csp = 'Microsoft Enhanced RSA and AES Cryptographic Provider\0'.encode('utf-16-le')
    header = struct.pack('<IIIIIIII', 0x24, 0, AES_ALG[key_bits], 0x8004, key_bits, 0x18, 0, 0) + csp
    info = struct.pack('<HHII', 4, 2, 0x24, len(header)) + header
    info += struct.pack('<I', 16) + salt + aes('ecb', key, verifier) + struct.pack('<I', 20) + aes('ecb', key, hashlib.sha1(verifier).digest() + bytes(12))
    package = struct.pack('<Q', len(plain)) + aes('ecb', key, plain + bytes(-len(plain) % 16))
    return cfb_write({'EncryptionInfo': info, 'EncryptedPackage': package})


# ------------------------------------------ xls and doc: XOR, RC4, CryptoAPI

def narrow(password, codec=None):
    """One byte for each UTF-16 unit of the password: the low byte, or the high one where the low one is zero.
    With `codec` the password is written in that code page instead, as Excel does on a system that has it."""
    if codec:
        return password.decode('utf-16-le').encode(codec)
    out = bytearray()
    for (unit,) in struct.iter_unpack('<H', password):
        out.append(unit & 0xFF or unit >> 8)
    return bytes(out)


def rotl(value, bits, width=8):
    mask = (1 << width) - 1
    return ((value << bits) | (value >> (width - bits))) & mask


def xor_key_and_hash(chars):
    """The 16-bit key and the 16-bit verifier of the XOR obfuscation, the way LibreOffice computes them (no tables)."""
    key, base, end = 0, 0x8000, 0xFFFF
    for char in reversed(chars):
        char &= 0x7F
        for _ in range(8):
            base = rotl(base, 1, 16)
            if base & 1:
                base ^= 0x1020
            if char & 1:
                key ^= base
            char >>= 1
            end = rotl(end, 1, 16)
            if end & 1:
                end ^= 0x1020
    key ^= end
    verifier = len(chars)
    if chars:
        verifier ^= 0xCE4B
    for index, char in enumerate(chars):
        verifier ^= rotl(char, (index + 1) % 15, 15)
    return (key if chars else 0), verifier


FILL = bytes.fromhex('bbffffbaffffb98000be0f00bf0f00')


def xor_array(chars, rotate):
    key, _ = xor_key_and_hash(chars)
    raw = bytearray(chars[:16] + FILL[:16 - len(chars)])
    for i in range(16):
        raw[i] = rotl(raw[i] ^ ((key >> 8) if i & 1 else (key & 0xFF)), rotate)
    return bytes(raw)


def rc4_97_key(password, salt, block):
    """Office 97 RC4: 40 bits of the password's MD5, stretched with the salt."""
    first = hashlib.md5(password).digest()[:5]
    second = hashlib.md5((first + salt) * 16).digest()[:5]
    return hashlib.md5(second + struct.pack('<I', block)).digest()


def capi_key(password, salt, block, key_bits):
    h = hashlib.sha1(hashlib.sha1(salt + password).digest() + struct.pack('<I', block)).digest()
    return h[:5] + bytes(11) if key_bits == 40 else h[:key_bits // 8]


def rc4_97_header(password, seed):
    salt = hashlib.shake_256(seed + b'salt').digest(16)
    verifier = hashlib.shake_256(seed + b'verifier').digest(16)
    sealed = rc4(rc4_97_key(password, salt, 0), verifier + hashlib.md5(verifier).digest())
    return struct.pack('<HH', 1, 1) + salt + sealed, lambda block: rc4_97_key(password, salt, block)


def capi_header(password, seed, key_bits):
    salt = hashlib.shake_256(seed + b'salt').digest(16)
    verifier = hashlib.shake_256(seed + b'verifier').digest(16)
    csp = 'Microsoft Enhanced Cryptographic Provider v1.0\0'.encode('utf-16-le')
    header = struct.pack('<IIIIIIII', 0x04, 0, 0x6801, 0x8004, key_bits, 1, 0, 0) + csp
    # A key size of 0 in the header stands for 40 bits.
    key_bits = key_bits or 40
    sealed = rc4(capi_key(password, salt, 0, key_bits), verifier + hashlib.sha1(verifier).digest())
    out = struct.pack('<HHII', 3, 2, 0x04, len(header)) + header + struct.pack('<I', 16) + salt + sealed[:16] + struct.pack('<I', 20) + sealed[16:]
    return out, lambda block: capi_key(password, salt, block, key_bits)


def rc4_stream(data, key_of, block_size, plain_ranges):
    """Encrypts a whole stream block by block; the bytes of `plain_ranges` stay as they are but use up key stream."""
    out = bytearray()
    for start in range(0, len(data), block_size):
        out += rc4(key_of(start // block_size), data[start:start + block_size])
    for a, b in plain_ranges:
        out[a:b] = data[a:b]
    return bytes(out)


XLS_PLAIN = {0x0809, 0x002F, 0x0194, 0x0195, 0x00E1, 0x0196, 0x0138}


def xls_encrypt(stream, password, method, seed, key_bits, codec=None):
    """A Workbook stream with a FILEPASS record put in after the first BOF, and everything after it encrypted."""
    if method == 'xor':
        chars = narrow(password, codec)
        key, verifier = xor_key_and_hash(chars)
        # Before BIFF8 the record is the key and the verifier alone; BIFF8 says which scheme first.
        biff8 = struct.unpack_from('<H', stream, 4)[0] == 0x0600
        body = struct.pack('<HHH', 0, key, verifier) if biff8 else struct.pack('<HH', key, verifier)
    elif method == 'rc4':
        header, key_of = rc4_97_header(password, seed)
        body = struct.pack('<H', 1) + header
    else:
        header, key_of = capi_header(password, seed, key_bits)
        body = struct.pack('<H', 1) + header
    filepass = struct.pack('<HH', 0x002F, len(body)) + body
    bof_end = 4 + struct.unpack_from('<H', stream, 2)[0]
    moved = bytearray(stream[:bof_end] + filepass + stream[bof_end:])
    # The sheets are found through BOUNDSHEET, which holds their positions in the stream.
    plain_ranges, records, at = [], [], 0
    while at + 4 <= len(moved):
        rid, size = struct.unpack_from('<HH', moved, at)
        if rid == 0x0085 and size >= 4 and at > bof_end:
            struct.pack_into('<I', moved, at + 4, struct.unpack_from('<I', moved, at + 4)[0] + len(filepass))
        records.append((rid, at + 4, size))
        plain_ranges.append((at, at + 4))
        if rid in XLS_PLAIN:
            plain_ranges.append((at + 4, at + 4 + size))
        elif rid == 0x0085:
            plain_ranges.append((at + 4, at + 8))
        at += 4 + size
    if method != 'xor':
        return rc4_stream(bytes(moved), key_of, 1024, plain_ranges)
    array = xor_array(chars, 2)
    out = bytearray(moved)
    for rid, start, size in records:
        if rid in XLS_PLAIN:
            continue
        skip = 4 if rid == 0x0085 else 0
        for i in range(skip, size):
            # Encrypting is decrypting backwards: XOR with the key, then rotate right by three (left by five).
            out[start + i] = rotl(moved[start + i] ^ array[(start + size + i) & 15], 5)
    return bytes(out)


def doc_encrypt(streams, password, method, seed, key_bits, codec=None):
    word = bytearray(streams['WordDocument'])
    flags = struct.unpack_from('<H', word, 10)[0]
    table_name = '1Table' if flags & 0x0200 else '0Table'
    table = streams[table_name]
    out = dict(streams)
    if method == 'xor':
        chars = narrow(password, codec)
        key, verifier = xor_key_and_hash(chars)
        array = xor_array(chars, 7)
        struct.pack_into('<H', word, 10, flags | 0x0100 | 0x8000)
        struct.pack_into('<HH', word, 14, verifier, key)

        def obscure(data, start):
            res = bytearray(data)
            for i in range(start, len(data)):
                k = array[i & 15]
                # A zero byte and a byte equal to its key byte are left alone (they would come out as each other).
                if data[i] and data[i] != k:
                    res[i] = data[i] ^ k
            return bytes(res)

        out['WordDocument'] = obscure(bytes(word), 68)
        out[table_name] = obscure(table, 0)
        if 'Data' in streams:
            out['Data'] = obscure(streams['Data'], 0)
        return out
    header, key_of = rc4_97_header(password, seed) if method == 'rc4' else capi_header(password, seed, key_bits)
    # The header goes in front of the table stream: everything the FIB points at there moves by its size.
    pairs = struct.unpack_from('<H', word, 152)[0]
    for i in range(pairs):
        fc, lcb = struct.unpack_from('<II', word, 154 + i * 8)
        if lcb:
            struct.pack_into('<I', word, 154 + i * 8, fc + len(header))
    struct.pack_into('<H', word, 10, (flags | 0x0100) & ~0x8000)
    struct.pack_into('<I', word, 14, len(header))
    out['WordDocument'] = rc4_stream(bytes(word), key_of, 512, [(0, 68)])
    out[table_name] = rc4_stream(header + table, key_of, 512, [(0, len(header))])
    if 'Data' in streams:
        out['Data'] = rc4_stream(streams['Data'], key_of, 512, [])
    return out


# ------------------------------------------------------------------- commands

def main(argv):
    command = argv[1]
    opts = dict(zip(argv[2::2], argv[3::2]))
    password = bytes.fromhex(opts.get('--password', ''))
    seed = opts.get('--seed', 'seed').encode('utf-8')
    key_bits = int(opts.get('--bits', '256'))
    data = sys.stdin.buffer.read()
    out = sys.stdout.buffer
    if command == 'decrypt':
        streams, _, _ = cfb_read(data)
        major, minor = struct.unpack_from('<HH', streams['EncryptionInfo'], 0)
        plain, hmac_ok = agile_decrypt(streams, password) if (major, minor) == (4, 4) else standard_decrypt(streams, password)
        if hmac_ok is False:
            raise Failure(BAD_HMAC, 'the integrity check failed')
        if opts.get('--zip') == 'check':
            try:
                with zipfile.ZipFile(io.BytesIO(plain)) as z:
                    broken = z.testzip()
            except zipfile.BadZipFile as error:
                raise Failure(BAD_ZIP, str(error))
            if broken is not None:
                raise Failure(BAD_ZIP, 'broken entry: ' + broken)
        sys.stderr.write('kind=%s hmac=%s size=%d\n' % ('agile' if (major, minor) == (4, 4) else 'standard', hmac_ok, len(plain)))
        out.write(plain)
    elif command == 'agile-encrypt':
        out.write(agile_encrypt(data, password, opts.get('--hash', 'SHA512'), key_bits, int(opts.get('--spin', '100000')), seed, int(opts.get('--password-bits', '0'))))
    elif command == 'standard-encrypt':
        out.write(standard_encrypt(data, password, key_bits, seed))
    elif command == 'agile-key':
        # The key of one block key for a salt: a known answer for the tests.
        base = spin(PY_HASH[opts['--hash']], bytes.fromhex(opts['--salt']), password, int(opts['--spin']))
        key = fit(hashlib.new(PY_HASH[opts['--hash']], base + bytes.fromhex(opts['--block'])).digest(), key_bits // 8)
        out.write(key.hex().encode('ascii'))
    elif command == 'standard-key':
        out.write(standard_key(bytes.fromhex(opts['--salt']), password, key_bits // 8).hex().encode('ascii'))
    elif command == 'cfb-list':
        streams, storages, facts = cfb_read(data, strict=True)
        listing = dict(facts, storages=storages, streams={path: dict(size=len(body), sha256=hashlib.sha256(body).hexdigest()) for path, body in streams.items()})
        out.write(json.dumps(listing).encode('utf-8'))
    elif command == 'cfb-stream':
        streams, _, _ = cfb_read(data)
        out.write(streams[opts['--path']])
    elif command == 'cfb-write':
        # {name: hex} on standard input.
        out.write(cfb_write({name: bytes.fromhex(body) for name, body in json.loads(data.decode('utf-8')).items()}))
    elif command == 'xls-encrypt':
        out.write(xls_encrypt(data, password, opts['--method'], seed, int(opts.get('--bits', '128')), opts.get('--codepage')))
    elif command == 'doc-encrypt':
        streams, _, _ = cfb_read(data)
        out.write(cfb_write(doc_encrypt(streams, password, opts['--method'], seed, int(opts.get('--bits', '128')), opts.get('--codepage'))))
    elif command == 'xor-key':
        chars = narrow(password)
        key, verifier = xor_key_and_hash(chars)
        out.write(json.dumps(dict(key=key, verifier=verifier, xls=xor_array(chars, 2).hex(), doc=xor_array(chars, 7).hex())).encode('utf-8'))
    else:
        raise Failure(OTHER, 'unknown command: ' + command)


if __name__ == '__main__':
    try:
        main(sys.argv)
    except Failure as failure:
        sys.stderr.write(str(failure) + '\n')
        sys.exit(failure.code)
    except Exception as error:  # noqa: BLE001 - whatever went wrong is the caller's to read
        sys.stderr.write('%s: %s\n' % (type(error).__name__, error))
        sys.exit(OTHER)
