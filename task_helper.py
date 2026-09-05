# 오늘 할 일 3개를 물어보고 저장해요.
task1 = input("오늘 할 첫 번째 일은 무엇인가요? ")
task2 = input("오늘 할 두 번째 일은 무엇인가요? ")
task3 = input("오늘 할 세 번째 일은 무엇인가요? ")

# 가장 중요한 할 일의 번호를 물어봐요.
important = input("가장 중요한 할 일은 몇 번인가요? (1, 2, 3): ")

# 올바른 번호를 입력할 때까지 다시 물어봐요.
while important not in ("1", "2", "3"):
    print("1, 2, 3 중 하나를 입력해 주세요.")
    important = input("가장 중요한 할 일은 몇 번인가요? (1, 2, 3): ")

# 고른 할 일에 표시를 붙여요.
if important == "1":
    task1 = task1 + " [가장 중요]"
elif important == "2":
    task2 = task2 + " [가장 중요]"
else:
    task3 = task3 + " [가장 중요]"

# 입력한 할 일을 번호와 함께 보여줘요.
print("\n오늘 할 일 목록")
print("1.", task1)
print("2.", task2)
print("3.", task3)

print("하나씩 천천히 해봐요!")
