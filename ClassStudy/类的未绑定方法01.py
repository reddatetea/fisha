class BaseClass:          #父类
    def foo(self):
        print('父类中定义的foo方法')

class SubClass(BaseClass):    #子类
    def foo(self):
        print('子类重写父类中的foo方法')
    def bar(self):
        print('执行bar方法')
        self.foo()
        BaseClass.foo(self)                   #使用类名调用实例主法（未绑定方法）调用父类被重写的方法
sc = SubClass()
sc.bar()