import unittest

def fix_mall(value: str, add_prefix: bool):
                if not value:
                    return ''
                prefixes = ("ТЦ", "ТРЦ", "ТРК", "ТД", "ТК")
                if any(value.startswith(prefix) for prefix in prefixes):
                    return value
                return f"ТЦ {value}" if add_prefix else value

class TestFixMall(unittest.TestCase):
    def test_fix_mall(self):
            testcases = [
                {
                    "name": "already_sorted",
                    "value": "Привет",
                    "add_prefix": True,
                    "expected": "ТЦ Привет",
                },
                {
                    "name": "already_sorted",
                    "value": "ТЦ Привет",
                    "add_prefix": True,
                    "expected": "ТЦ Привет",
                },
                {
                    "name": "already_sorted",
                    "value": "ТРК \"Привет\"",
                    "add_prefix": True,
                    "expected": "ТРК \"Привет\"",
                },
            ]

            for case in testcases:
                actual = fix_mall(case["value"], case["add_prefix"])
                self.assertEqual(
                    case["expected"],
                    actual,
                    "failed test {} expected {}, actual {}".format(
                        case["name"], case["expected"], actual
                    ),
                )