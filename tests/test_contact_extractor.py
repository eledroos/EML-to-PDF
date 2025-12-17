"""Tests for contact extraction and address book generation."""

import csv
import tempfile
from pathlib import Path

import pytest

from eml_to_pdf.contact_extractor import (
    Contact,
    parse_email_header,
    extract_contacts_from_metadata,
    deduplicate_contacts,
    generate_address_book,
)


class TestParseEmailHeader:
    """Tests for parse_email_header function."""

    def test_simple_email(self):
        """Test parsing a simple email address."""
        result = parse_email_header("john@example.com")
        assert len(result) == 1
        assert result[0] == ("john", "john@example.com")

    def test_email_with_name(self):
        """Test parsing 'Name <email>' format."""
        result = parse_email_header("John Doe <john@example.com>")
        assert len(result) == 1
        assert result[0] == ("John Doe", "john@example.com")

    def test_multiple_addresses(self):
        """Test parsing comma-separated addresses."""
        result = parse_email_header("John <john@example.com>, Jane <jane@example.com>")
        assert len(result) == 2
        assert result[0] == ("John", "john@example.com")
        assert result[1] == ("Jane", "jane@example.com")

    def test_empty_value(self):
        """Test parsing empty or placeholder values."""
        assert parse_email_header("") == []
        assert parse_email_header("No Recipients") == []
        assert parse_email_header("No CC") == []
        assert parse_email_header("No BCC") == []
        assert parse_email_header("Unknown Sender") == []

    def test_email_lowercase(self):
        """Test that emails are normalized to lowercase."""
        result = parse_email_header("John@EXAMPLE.COM")
        assert result[0][1] == "john@example.com"

    def test_mixed_formats(self):
        """Test parsing mixed format addresses."""
        result = parse_email_header("plain@example.com, Named User <named@example.com>")
        assert len(result) == 2
        assert result[0] == ("plain", "plain@example.com")
        assert result[1] == ("Named User", "named@example.com")


class TestExtractContactsFromMetadata:
    """Tests for extract_contacts_from_metadata function."""

    def test_basic_extraction(self):
        """Test extracting contacts from metadata dict."""
        metadata = {
            "sender": "sender@example.com",
            "recipients": "recipient@example.com",
            "cc": "cc@example.com",
            "bcc": "bcc@example.com",
        }
        contacts = extract_contacts_from_metadata(metadata)

        assert len(contacts) == 4
        types = [c.contact_type for c in contacts]
        assert "From" in types
        assert "To" in types
        assert "CC" in types
        assert "BCC" in types

    def test_empty_fields(self):
        """Test handling of empty/placeholder fields."""
        metadata = {
            "sender": "sender@example.com",
            "recipients": "No Recipients",
            "cc": "",
            "bcc": "No BCC",
        }
        contacts = extract_contacts_from_metadata(metadata)

        assert len(contacts) == 1
        assert contacts[0].contact_type == "From"

    def test_multiple_recipients(self):
        """Test extracting multiple recipients."""
        metadata = {
            "sender": "sender@example.com",
            "recipients": "alice@example.com, bob@example.com",
            "cc": "",
            "bcc": "",
        }
        contacts = extract_contacts_from_metadata(metadata)

        assert len(contacts) == 3
        to_contacts = [c for c in contacts if c.contact_type == "To"]
        assert len(to_contacts) == 2


class TestDeduplicateContacts:
    """Tests for deduplicate_contacts function."""

    def test_removes_duplicates(self):
        """Test that duplicates are removed by email."""
        contacts = [
            Contact(name="John", email="john@example.com", contact_type="From"),
            Contact(name="John Doe", email="john@example.com", contact_type="To"),
            Contact(name="Jane", email="jane@example.com", contact_type="To"),
        ]
        result = deduplicate_contacts(contacts)

        assert len(result) == 2
        emails = [c.email for c in result]
        assert "john@example.com" in emails
        assert "jane@example.com" in emails

    def test_case_insensitive(self):
        """Test that deduplication is case-insensitive."""
        contacts = [
            Contact(name="John", email="JOHN@example.com", contact_type="From"),
            Contact(name="John", email="john@example.com", contact_type="To"),
        ]
        result = deduplicate_contacts(contacts)

        assert len(result) == 1

    def test_keeps_first_occurrence(self):
        """Test that first occurrence is kept."""
        contacts = [
            Contact(name="First Name", email="test@example.com", contact_type="From"),
            Contact(name="Second Name", email="test@example.com", contact_type="To"),
        ]
        result = deduplicate_contacts(contacts)

        assert len(result) == 1
        assert result[0].name == "First Name"
        assert result[0].contact_type == "From"

    def test_empty_list(self):
        """Test deduplicating empty list."""
        assert deduplicate_contacts([]) == []


class TestGenerateAddressBook:
    """Tests for generate_address_book function."""

    def test_creates_csv(self):
        """Test that CSV file is created with correct format."""
        contacts = [
            Contact(name="Alice", email="alice@example.com", contact_type="From"),
            Contact(name="Bob", email="bob@example.com", contact_type="To"),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "address_book.csv"
            result = generate_address_book(contacts, output_path)

            assert result == output_path
            assert output_path.exists()

            # Verify CSV contents
            with open(output_path, "r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

            assert rows[0] == ["Name", "Email", "Type"]
            assert len(rows) == 3  # header + 2 contacts

    def test_sorted_by_name(self):
        """Test that contacts are sorted alphabetically by name."""
        contacts = [
            Contact(name="Zack", email="zack@example.com", contact_type="To"),
            Contact(name="Alice", email="alice@example.com", contact_type="From"),
            Contact(name="Bob", email="bob@example.com", contact_type="CC"),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "address_book.csv"
            generate_address_book(contacts, output_path, dedupe=False)

            with open(output_path, "r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

            # Skip header, check order
            assert rows[1][0] == "Alice"
            assert rows[2][0] == "Bob"
            assert rows[3][0] == "Zack"

    def test_empty_contacts(self):
        """Test handling of empty contacts list."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "address_book.csv"
            result = generate_address_book([], output_path)

            assert result is None
            assert not output_path.exists()

    def test_dedupe_option(self):
        """Test that dedupe parameter works."""
        contacts = [
            Contact(name="John", email="john@example.com", contact_type="From"),
            Contact(name="John D", email="john@example.com", contact_type="To"),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            # With dedupe (default)
            path1 = Path(tmpdir) / "with_dedupe.csv"
            generate_address_book(contacts, path1, dedupe=True)

            with open(path1, "r", newline="", encoding="utf-8") as f:
                rows1 = list(csv.reader(f))
            assert len(rows1) == 2  # header + 1 contact

            # Without dedupe
            path2 = Path(tmpdir) / "without_dedupe.csv"
            generate_address_book(contacts, path2, dedupe=False)

            with open(path2, "r", newline="", encoding="utf-8") as f:
                rows2 = list(csv.reader(f))
            assert len(rows2) == 3  # header + 2 contacts
