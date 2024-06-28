from datetime import datetime


def log_print(*objects):
    with open('log.txt', 'a') as f:
        print(datetime.now(), *objects, file=f)
        #print(datetime.now(), *objects)


if __name__ == '__main__':
    log_print('hello world')
    a = 1
    b = 'str'
    log_print(a, b)
    log_print(f'dsadsa{a}dsadsa')