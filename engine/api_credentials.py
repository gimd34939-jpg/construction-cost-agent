"""Per-Windows-user encrypted credentials. Secret values never enter job artifacts."""
import argparse
import base64
import ctypes
from ctypes import wintypes
import getpass
import json
import os
from pathlib import Path
import re
import sys
from datetime import datetime, timezone

ENTROPY = b'CodexConstructionCost/api/v1'
STORE = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData/Local'))) / 'CodexConstructionCost' / 'api_credentials.dpapi.json'
ENV_NAMES = {'PPS': 'DATA_GO_KR_SERVICE_KEY', 'OPINET': 'OPINET_API_KEY'}

class _Blob(ctypes.Structure):
    _fields_ = [('cbData', wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_byte))]

def _blob(value):
    buffer = ctypes.create_string_buffer(value)
    return _Blob(len(value), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer

def _crypt(value, decrypt=False):
    if os.name != 'nt':
        raise RuntimeError('Windows DPAPI required for protected credential storage')
    incoming, keep_in = _blob(value)
    entropy, keep_entropy = _blob(ENTROPY)
    output = _Blob()
    dll = ctypes.WinDLL('crypt32', use_last_error=True)
    if decrypt:
        ok = dll.CryptUnprotectData(ctypes.byref(incoming), None, ctypes.byref(entropy), None, None, 1, ctypes.byref(output))
    else:
        ok = dll.CryptProtectData(ctypes.byref(incoming), 'Construction API credential', ctypes.byref(entropy), None, None, 1, ctypes.byref(output))
    if not ok:
        raise RuntimeError('Credential encryption/decryption failed')
    try:
        return ctypes.string_at(output.pbData, output.cbData)
    finally:
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.LocalFree.argtypes = [ctypes.c_void_p]
        kernel.LocalFree.restype = ctypes.c_void_p
        kernel.LocalFree(output.pbData)

def _read():
    if not STORE.exists():
        return {'version': 1, 'secrets': {}}
    return json.loads(STORE.read_text(encoding='utf-8-sig'))

def get_secret(provider):
    provider = provider.upper()
    if provider not in ENV_NAMES:
        raise ValueError('Unsupported credential provider')
    if os.environ.get(ENV_NAMES[provider]):
        return os.environ[ENV_NAMES[provider]].strip()
    saved = _read().get('secrets', {}).get(provider)
    if not saved:
        raise RuntimeError(f'{provider} credential is not registered')
    try:
        return _crypt(base64.b64decode(saved['ciphertext']), decrypt=True).decode('utf-8')
    except Exception:
        raise RuntimeError(f'{provider} protected credential cannot be read by this Windows user') from None

def set_secret(provider, value, source=None):
    provider = provider.upper()
    value = value.strip()
    if provider not in ENV_NAMES or not value or any(c.isspace() for c in value):
        raise ValueError('Invalid credential value or provider')
    if provider == 'PPS' and not re.fullmatch(r'[0-9a-fA-F]{64}|[A-Za-z0-9+/=%]{40,}', value):
        raise ValueError('Invalid PPS credential format')
    data = _read()
    data['secrets'][provider] = {'ciphertext': base64.b64encode(_crypt(value.encode('utf-8'))).decode('ascii'), 'registered_at': datetime.now(timezone.utc).isoformat(), 'source': source or {'type':'USER_REGISTRATION'}}
    STORE.parent.mkdir(parents=True, exist_ok=True)
    temporary = STORE.with_suffix('.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, STORE)
    return {'provider': provider, 'registered': True, 'protection':'WINDOWS_CURRENT_USER_DPAPI'}

def credential_status():
    result = {}
    for provider in ENV_NAMES:
        try:
            result[provider] = {'registered':bool(get_secret(provider)), 'protection':'ENVIRONMENT' if os.environ.get(ENV_NAMES[provider]) else 'WINDOWS_CURRENT_USER_DPAPI'}
        except Exception:
            result[provider] = {'registered':False}
    return result

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['status', 'register', 'stdin-register'])
    parser.add_argument('--provider', choices=['PPS','OPINET'])
    args = parser.parse_args()
    if args.action == 'status':
        result = credential_status()
    elif args.action == 'stdin-register':
        payload = json.load(sys.stdin)
        result = [set_secret(p, v, {'type':'USER_SUPPLIED_SCREENSHOT'}) for p,v in payload.items()]
    else:
        if not args.provider:
            parser.error('--provider required')
        result = set_secret(args.provider, getpass.getpass(f'{args.provider} API key (hidden): '))
    print(json.dumps(result, ensure_ascii=False))

if __name__ == '__main__':
    main()
