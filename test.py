import asyncio

from easy_access.pdf.parse import test_func

result = asyncio.run(test_func())
print(result)
