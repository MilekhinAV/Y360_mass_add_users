import importlib.util
import sys
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).with_name("import_yandex360_users.py")
SPEC = importlib.util.spec_from_file_location("yandex360_import", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class FakeClient:
    def __init__(self, departments=None, users=None):
        self.departments = departments or [
            {"id": 1, "parentId": 0, "name": "Все сотрудники"}
        ]
        self.users = users or []
        self.created_departments = []
        self.created_users = []
        self.next_department_id = 100

    def list_departments(self):
        return list(self.departments)

    def list_users(self):
        return list(self.users)

    def create_department(self, name, parent_id):
        item = {"id": self.next_department_id, "parentId": parent_id, "name": name}
        self.next_department_id += 1
        self.created_departments.append(item)
        self.departments.append(item)
        return item

    def create_user(self, payload):
        self.created_users.append(payload)
        return {"id": str(1000 + len(self.created_users)), **payload}


def example_row(path=("Школы", "Школа №1")):
    return MODULE.UserRow(
        row_number=2,
        last_name="Иванов",
        first_name="Иван",
        middle_name="Иванович",
        login="ivanovii",
        password="",
        must_change_password=True,
        position="Специалист",
        gender="male",
        birthday="",
        language="",
        work_phone="89111111111",
        mobile_phone="",
        personal_email="mail@ya.ru",
        department_path=path,
    )


class ImportTests(unittest.TestCase):
    def test_personal_email_is_contact_not_about(self):
        payload = MODULE.user_payload(example_row(), 42)
        self.assertNotIn("about", payload)
        self.assertIn(
            {"type": "email", "value": "mail@ya.ru", "label": "Personal"},
            payload["contacts"],
        )

    def test_dry_run_plans_hierarchy_without_post(self):
        client = FakeClient()
        summary = MODULE.run_import([example_row()], client=client, dry_run=True)
        self.assertEqual(summary.departments_planned, 2)
        self.assertEqual(summary.users_planned, 1)
        self.assertEqual(client.created_departments, [])
        self.assertEqual(client.created_users, [])

    def test_real_run_creates_hierarchy_and_user(self):
        client = FakeClient()
        summary = MODULE.run_import([example_row()], client=client, dry_run=False)
        self.assertEqual(summary.departments_created, 2)
        self.assertEqual(summary.users_created, 1)
        self.assertEqual(client.created_users[0]["departmentId"], 101)

    def test_existing_branch_is_reused(self):
        client = FakeClient(
            departments=[
                {"id": 1, "parentId": 0, "name": "Все сотрудники"},
                {"id": 10, "parentId": 1, "name": "Школы"},
            ]
        )
        summary = MODULE.run_import([example_row()], client=client, dry_run=False)
        self.assertEqual(summary.departments_created, 1)
        self.assertEqual(client.created_departments[0]["parentId"], 10)

    def test_existing_login_is_skipped(self):
        client = FakeClient(users=[{"id": "77", "nickname": "IvanovII"}])
        summary = MODULE.run_import([example_row()], client=client, dry_run=False)
        self.assertEqual(summary.users_skipped, 1)
        self.assertEqual(client.created_users, [])
        self.assertEqual(client.created_departments, [])


if __name__ == "__main__":
    unittest.main()
