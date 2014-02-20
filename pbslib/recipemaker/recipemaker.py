""" Simple recipe maker module.
"""


class Recipe():
    """ Recipe maker.

    @ivar buffer: line-buffer containing list of steps.
    @type buffer: list[str]
    @ivar Name: recipe name(optional)
    @type Name: str
    """

    def __init__(self, Name=''):
        self.Name = Name
        self.buffer = ['']  # initialize with empty line.

    def write(self, code):
        """
        @param code: code to write
        @type code: str
        """
        self.buffer.append(code)

    def set(self, param, value):
        """
        @param param: variable to set
        @type param: str | RecipeVariable
        @param value: value to set it to
        @type value: str | int | float
        """
        self.write("Set \"%s\" to %s" % (param, str(value)))

    def clear(self):
        """
        @return: None
        @rtype: None
        Clear the recipe.
        """
        self.buffer.clear()
        self.buffer.append('')

    # alias
    Set = set

    def waituntil(self, param, op, val):
        """
        @param param: parameter to wait for
        @type param: str
        @param op: wait operation
        @type op: str
        @param val: value to wait for
        @type val: str | int | float
        """
        self.write("Wait until \"%s\" %s %s" % (param, op, str(val)))

    wait_until = waituntil

    def wait(self, sec):
        """
        @param sec: seconds to wait
        @type sec: int
        """
        sec = int(sec)
        self.write("Wait %d seconds" % sec)

    def getvalue(self):
        return str(self)

    def __str__(self):
        return '\n'.join(self.buffer)

    def __len__(self):
        return len(self.buffer)

    import sys

    def print(self, stream=sys.stdout):
        print(str(self), file=stream)


class LongRecipe(Recipe):
    """ Chain together multiple recipes
    """

    def add_recipe(self, recipe):
        """
        @param recipe: recipe to add
        @type recipe: Recipe
        @return: None
        @rtype: None
        """
        self.buffer.extend(recipe.buffer)

    def extend_recipes(self, recipe_list):
        """
        @param recipe_list: list of recipes to add
        @type recipe_list: collections.Iterable[Recipe]
        @return: None
        @rtype: None
        """
        for recipe in recipe_list:
            self.buffer.extend(recipe.buffer)


class RecipeVariable():
    """
    Create an object that returns a string
    corresponding to a recipe step for
    "wait until" steps.

    @type _name: str
    """

    # Double brackets escape the formatting from the cls template
    # so they become valid .format() args in instance template.
    wait_step_template = 'Wait until "{var}" {{cmp}} {{val}}'
    set_step_template = 'Set "{var}" to {{val}}'

    def __init__(self, name):
        """
        @param name: name of property
        @type name: str
        @return: None
        @rtype: None
        """
        self._name = name
        self.wait_step_template = self.wait_step_template.format(var=name)
        self.set_step_template = self.set_step_template.format(var=name)

    @property
    def Name(self):
        """
         @rtype: str
        """
        return self._name

    name = Name

    def wait_until(self, cmp, val):
        """
        @param cmp: comparison operator
        @type cmp: str
        @param val: value to wait for
        @type val: str | float | int
        @rtype: str
        """

        if cmp not in ("<", ">", "=", "<=", ">=", "!="):
            raise ValueError("'%s' is an invalid recipe comparison operator" % cmp)

        return self.wait_step_template.format(
                                        var=self._name,
                                        cmp=cmp,
                                        val=val)

    def set(self, val):
        """
        @type val: str | float | int
        """

        return self.set_step_template.format(val=val)

    def __lt__(self, val):
        """ < """

        return self.wait_until("<", val)

    def __le__(self, val):
        """ <= """
        return self.wait_until("<=", val)

    def __eq__(self, val):
        """ == """
        return self.wait_until("=", val)

    def __ne__(self, val):
        """ != """
        return self.wait_until("!=", val)

    def __gt__(self, val):
        """ > """
        return self.wait_until(">", val)

    def __ge__(self, val):
        """ >= """
        return self.wait_until(">=", val)

    def __str__(self):
        return self._name


if __name__ == '__main__':
    r = RecipeVariable("foo")
    print(r > 5)
    r2 = Recipe()
    r2.set(r, 5)
    print(r2)
    print(r.set(5))
