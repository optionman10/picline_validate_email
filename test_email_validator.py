"""Unit tests for email_validator core logic (no GUI, no network)."""

import unittest
from unittest.mock import patch

from email_validator import (
    DomainChecker,
    Severity,
    Status,
    validate_email,
)


def no_dns():
    """A DomainChecker that never does network I/O."""
    return DomainChecker(enabled=False)


class TestSyntax(unittest.TestCase):
    def check(self, email, **kw):
        return validate_email(email, no_dns())

    def test_valid_simple(self):
        r = self.check("john.doe@example.com")
        self.assertEqual(r.status, Status.VALID)
        self.assertEqual(r.error_description, "OK")

    def test_valid_plus_tag(self):
        self.assertEqual(self.check("user+tag@gmail.com").status, Status.VALID)

    def test_missing_at(self):
        r = self.check("not-an-email.com")
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("exactly one '@'", r.error_description)

    def test_double_at(self):
        r = self.check("a@b@c.com")
        self.assertEqual(r.status, Status.INVALID)

    def test_empty(self):
        self.assertEqual(self.check("").status, Status.INVALID)
        self.assertEqual(self.check(None).status, Status.INVALID)
        self.assertEqual(self.check(float("nan")).status, Status.INVALID)

    def test_whitespace_trimmed_but_warned(self):
        r = self.check("  john@example.com  ")
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)
        self.assertIn("whitespace", r.error_description.lower())

    def test_consecutive_dots(self):
        r = self.check("jo..hn@example.com")
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("consecutive dots", r.error_description)

    def test_leading_dot(self):
        r = self.check(".john@example.com")
        self.assertEqual(r.status, Status.INVALID)

    def test_trailing_dot(self):
        r = self.check("john.@example.com")
        self.assertEqual(r.status, Status.INVALID)

    def test_illegal_chars(self):
        r = self.check("jo hn@example.com")
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("Illegal character", r.error_description)

    def test_local_too_long(self):
        r = self.check("a" * 65 + "@example.com")
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("64", r.error_description)

    def test_bad_domain_no_tld(self):
        r = self.check("john@localhost")
        self.assertEqual(r.status, Status.INVALID)

    def test_bad_tld_digits(self):
        r = self.check("john@example.123")
        self.assertEqual(r.status, Status.INVALID)

    def test_uppercase_ok(self):
        self.assertEqual(self.check("John.Doe@Example.COM").status, Status.VALID)


class TestNamingConventions(unittest.TestCase):
    def check(self, email):
        return validate_email(email, no_dns())

    def test_generic_mailbox_warns(self):
        r = self.check("info@example.com")
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)
        self.assertIn("generic", r.error_description.lower())

    def test_noreply_warns(self):
        r = self.check("noreply@example.com")
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)

    def test_all_digits_warns(self):
        r = self.check("123456@example.com")
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)
        self.assertIn("digits", r.error_description.lower())

    def test_random_looking_warns(self):
        r = self.check("xk7qz9w2bp4@example.com")
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)
        self.assertIn("random", r.error_description.lower())

    def test_normal_name_no_warning(self):
        r = self.check("maria.garcia@example.com")
        self.assertEqual(r.status, Status.VALID)


class TestDomainChecks(unittest.TestCase):
    def test_disposable_domain_error(self):
        r = validate_email("someone@mailinator.com", no_dns())
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("disposable", r.error_description.lower())

    def test_typo_suggestion(self):
        r = validate_email("john@gmial.com", no_dns())
        self.assertEqual(r.status, Status.VALID_WITH_WARNINGS)
        self.assertEqual(r.suggestion, "john@gmail.com")
        self.assertIn("did you mean", r.error_description)

    def test_typo_suggestion_yahoo(self):
        r = validate_email("jane@yaho.com", no_dns())
        self.assertEqual(r.suggestion, "jane@yahoo.com")

    def test_no_suggestion_for_real_domain(self):
        r = validate_email("john@acmecorp.io", no_dns())
        self.assertIsNone(r.suggestion)

    def test_dns_no_mx_is_error(self):
        checker = DomainChecker(enabled=False)
        with patch.object(DomainChecker, "has_mx", return_value=False):
            r = validate_email("john@nonexistent-domain-xyz123.com", checker)
        self.assertEqual(r.status, Status.INVALID)
        self.assertIn("MX", r.error_description)

    def test_dns_unknown_is_not_penalized(self):
        checker = DomainChecker(enabled=False)
        with patch.object(DomainChecker, "has_mx", return_value=None):
            r = validate_email("john@example.com", checker)
        self.assertEqual(r.status, Status.VALID)

    def test_dns_ok(self):
        checker = DomainChecker(enabled=False)
        with patch.object(DomainChecker, "has_mx", return_value=True):
            r = validate_email("john@example.com", checker)
        self.assertEqual(r.status, Status.VALID)


class TestDomainCheckerCache(unittest.TestCase):
    def test_cache_hit_avoids_resolver(self):
        checker = DomainChecker(enabled=False)
        checker._cache["cached.com"] = True
        self.assertTrue(checker.has_mx("cached.com"))
        self.assertTrue(checker.has_mx("CACHED.COM"))  # case-insensitive

    def test_disabled_returns_none(self):
        checker = DomainChecker(enabled=False)
        self.assertIsNone(checker.has_mx("anything.com"))


class TestResultReporting(unittest.TestCase):
    def test_error_description_orders_errors_first(self):
        r = validate_email("info@gmial..com", no_dns())
        desc = r.error_description
        first_warn = desc.find("[WARNING]")
        first_err = desc.find("[ERROR]")
        self.assertNotEqual(first_err, -1)
        if first_warn != -1:
            self.assertLess(first_err, first_warn)


if __name__ == "__main__":
    unittest.main(verbosity=2)
