#           Copyright Alexander Malahov 2017-2018.
#  Distributed under the Boost Software License, Version 1.0.
#     (See accompanying file ../../LICENSE.txt or copy at
#           http://www.boost.org/LICENSE_1_0.txt)


def enum(enum_type_name, *values):
    return __import__('com.sun.star.' + enum_type_name, fromlist=list(values))
