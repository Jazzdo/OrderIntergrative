import random
RAND_MIN = 1
RAND_MAX = 10
answer = random.randint(RAND_MIN, RAND_MAX)

print(answer)
indata = int(input('{}에서 {} 사이의 수를 맞히세요 >> '.format(RAND_MIN,RAND_MAX)))
while True:

    if indata == answer:
        print('축하한다. {}: 정답이다.: '.format(indata))
        break
    if indata < answer:
        str = '{}보다 더 큰 수로 다시 입력 >> '.format(indata)
        indata = int(input(str))
    if indata > answer :
        str = '{}보다 더 작은 수로 다시 입력 >> '.format(indata)
        indata = int(input(str))
    if indata<RAND_MIN or indata>RAND_MAX:
        str = "이상한값: 다시 입력 >>"
        indata = int(input(str))

print("종료".center(30, '*'))