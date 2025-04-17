import unittest
from processing import ProcessPartA

class TestProcessPartA(unittest.TestCase):
    def setUp(self):
        """
        Set up the test environment with a ProcessPartA instance.
        """
        self.processor = ProcessPartA()

    # Includes valid values 
    def test_average_au_grade_standard(self):
        self.processor.data["samples"] = [
            {"Au": 0.3}, 
            {"Au": 0.5}, 
            {"Au": 1.0},  
            {"Au": 0.0}, 
            {"Au": -1.0} 
        ]
        result = self.processor.calculate_average_au_grade()
        expected = (0.3 + 0.5 + 1.0 + 0.0 + -1.0) / 5  
        self.assertAlmostEqual(result, expected)

    # Includes string value, missing key value
    def test_ignore_none_and_invalid_au(self):
        self.processor.data["samples"] = [
            {"Au": 1.0},
            {"Au": None},
            {"Au": "negative"},
            {"ID": "Missing"}
        ]
        result = self.processor.calculate_average_au_grade()
        expected = 1.0
        self.assertEqual(result, expected)
    
    # Includes all invalid values scenario
    def test_all_invalid_samples(self):
        self.processor.data["samples"] = [
            {"Au": None},
            {"Au": "NaN"},
            {"ID": "SAMP003"}
        ]
        result = self.processor.calculate_average_au_grade()
        self.assertEqual(result, 0.0)

    # Includes empty sample list
    def test_empty_sample_list(self):
        self.processor.data["samples"] = []
        self.assertEqual(self.processor.calculate_average_au_grade(), 0.0)

if __name__ == "__main__":
    unittest.main()
