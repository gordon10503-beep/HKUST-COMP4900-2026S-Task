#!/usr/bin/env python3
"""Generate a submission hash from student ID and ITSC email.

Students run this script locally, then copy the output values into Data/student.json.
"""

from __future__ import annotations

import hashlib
HASH_FUNCTION = "sha256"


def build_payload(student_id: str, itsc_email: str) -> str:
    # Use a canonical representation so every student hashes in the same way.
    return f"{student_id.strip()}|{itsc_email.strip().lower()}"


def compute_hash(payload: str) -> str:
    hasher = hashlib.new(HASH_FUNCTION)
    hasher.update(payload.encode("utf-8"))
    return hasher.hexdigest()


def main() -> None:
    print("=== COMP4900 Hash Generator ===")
    student_id = input("Enter your 8-digit student ID: ").strip()
    itsc_email = input("Enter your ITSC email (xxxx@connect.ust.hk): ").strip().lower()

    payload = build_payload(student_id, itsc_email)
    hash_value = compute_hash(payload)

    print("\nCopy this value into Data/student.json:")
    print(f"hash-value: {hash_value}")


if __name__ == "__main__":
    main()
