import unittest
from shuffle import divide_users


class TestShuffle(unittest.TestCase):
    def test_divide_users(self):
        users = [f'user{i}' for i in range(1, 54 + 1)]
        group = divide_users(users)

        user_set = set()
        for users in group:
            # 1当番3名か
            self.assertEqual(len(users), 3)

            # 1回のみ割り当てられているか
            for user in users:
                self.assertNotIn(user, user_set)
                user_set.add(user)


if __name__ == '__main__':
    unittest.main()
