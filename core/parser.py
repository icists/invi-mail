import json
import os


# Test file
test_file = "/Users/junss/tech/backend/mail/data/templates/spe_invi_ko.json"

class ParseError(Exception):
    pass

class TokenNotMatchError(ParseError):
    """
    Please check tokens. Tokens are not matched normally.
    """
    pass

class ContentParser():
    """
    @functionality
    - read content template
    - put args into template
    """
    
    def __init__(self, template : str, values : dict):
        # Parser Constants
        self.LTOKEN = '{'
        self.RTOKEN = '}'

        _file = open(template, 'r')
        _file = json.load(_file)
        self._title = _file['title']
        self._template = _file['template']

        self._values = values

        if self._is_valid_template():
            self._put_values()

    def get_title(self):

        _vars = self._values.keys()
        result = self._title
        for _tk in _vars:
            tk = '{' + _tk + '}'
            token_value = self._values[_tk]
            if token_value == None:
                token_value == ""
            result = result.replace(tk, token_value)
        return result

    def get_content(self):
        return '\n'.join(self._content)

    def _is_valid_template(self):
        """
        Check if the template is valid to put variables in, using stack.
        Do not support token parsing over two lines.
        """
        _test = self._template
        _check_stack = []
        for line in self._template:
            for c in line:
                if c == self.LTOKEN:
                    if len(_check_stack) != 0:
                        if _check_stack[-1] == self.LTOKEN:
                            raise TokenNotMatchError
                    _check_stack.append(c)
                if c == self.RTOKEN:
                    if _check_stack[-1] != self.LTOKEN:
                        raise TokenNotMatchError
                    _check_stack.pop()
        return True

    def _put_values(self):
        """
        put value at each variable
        """
        self._content = []
        _vars = self._values.keys()

        # Put value each line
        for line in self._template:
            # Iterate for every tokens
            _line = line
            for _tk in _vars:
                tk = '{' + _tk + '}'
                token_value = self._values[_tk]
                _line = _line.replace(tk, token_value)
            self._content.append(_line)
    
    def test(self):
        return "This is a test message"


# For testing.
if __name__ == "__main__":
    my_val = {"name" : "Bongjun", "age" : "21", "target" : "마이클 루덴"}
    ps = ContentParser(test_file, my_val)
    #print(ps._template)
    print(ps.get_content())
