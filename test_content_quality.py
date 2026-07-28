# -*- coding: utf-8 -*-
import unittest

from content_quality import normalize_raw
from to_schema_v3 import convert


class ContentQualityTest(unittest.TestCase):
    def sample(self):
        course = {
            "name": "중3 수학_도훈", "staffName": "도훈", "room": "A",
            "weeklySchedule": {"slots": [{"day": "MONDAY", "start": "17:00", "duration_minutes": 120}]},
        }
        return {
            "basic": {"name": "테스트학원", "subjects": {"math": True}},
            "courses": [course, dict(course), {
                "name": "class A",
                "weeklySchedule": {"slots": [{"day": "TUESDAY", "start": "14:00", "duration_minutes": 60}]},
            }],
            "content": {
                "achievements": '["30점에서 80점으로 향상", {"title":"내신","description":"4등급에서 2등급"}]',
                "specialPrograms": '[{"name":"방학 특강","description":"취약 단원 집중"}]',
                "admissionSteps": '["상담 신청", "진단 평가", "결과 상담", "반 배정", "수업 시작"]',
                "faq": '[{"question":"보강이 있나요?","answer":"규정에 따라 진행합니다."}]',
                "operatingRules": '{"absence":"결석 전 연락","withdrawal":"반복 미이행 시 상담"}',
            },
        }

    def test_deduplicates_only_identical_courses(self):
        normalized = normalize_raw(self.sample())
        self.assertEqual(len(normalized["courses"]), 2)
        self.assertEqual(normalized["_quality"]["counts"]["courseDuplicates"], 1)
        self.assertEqual(normalized["_quality"]["counts"]["admissionSteps"], 5)

    def test_schema_preserves_all_fact_items(self):
        schema = convert(self.sample())
        self.assertEqual(len(schema["achievements"]["items"]), 2)
        self.assertEqual(len(schema["admission"]["items"]), 5)
        self.assertEqual(len(schema["faq"]["items"]), 1)
        self.assertTrue(any(x["code"] == "POSSIBLE_PLACEHOLDER_COURSE"
                            for x in schema["_dataAudit"]["issues"]))


if __name__ == "__main__":
    unittest.main()
