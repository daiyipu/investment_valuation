from chapters.chapter01_overview import generate_chapter as ch01
from chapters.chapter02_financial import generate_chapter as ch02
from chapters.chapter03_industry import generate_chapter as ch03
from chapters.chapter04_dcf import generate_chapter as ch04
from chapters.chapter05_relative import generate_chapter as ch05
from chapters.chapter06_sensitivity import generate_chapter as ch06
from chapters.chapter07_montecarlo import generate_chapter as ch07
from chapters.chapter08_scenario import generate_chapter as ch08
from chapters.chapter09_stress import generate_chapter as ch09
from chapters.chapter10_var import generate_chapter as ch10
from chapters.chapter11_assessment import generate_chapter as ch11

# Module references for main.py
chapter01_overview = __import__('chapters.chapter01_overview', fromlist=['generate_chapter'])
chapter02_financial = __import__('chapters.chapter02_financial', fromlist=['generate_chapter'])
chapter03_industry = __import__('chapters.chapter03_industry', fromlist=['generate_chapter'])
chapter04_dcf = __import__('chapters.chapter04_dcf', fromlist=['generate_chapter'])
chapter05_relative = __import__('chapters.chapter05_relative', fromlist=['generate_chapter'])
chapter06_sensitivity = __import__('chapters.chapter06_sensitivity', fromlist=['generate_chapter'])
chapter07_montecarlo = __import__('chapters.chapter07_montecarlo', fromlist=['generate_chapter'])
chapter08_scenario = __import__('chapters.chapter08_scenario', fromlist=['generate_chapter'])
chapter09_stress = __import__('chapters.chapter09_stress', fromlist=['generate_chapter'])
chapter10_var = __import__('chapters.chapter10_var', fromlist=['generate_chapter'])
chapter11_assessment = __import__('chapters.chapter11_assessment', fromlist=['generate_chapter'])
