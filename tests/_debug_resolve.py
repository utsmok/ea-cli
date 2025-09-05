import asyncio
from tests.helpers import QuerySetMock
from easy_access.db.relations import _resolve_queryset_candidate

async def main():
    mock_pdf = type('P', (), {})()
    mock_pdf.material_id = 1001
    mock_pdf.replace_with = type('R', (), {'material_id':2001})()
    q = QuerySetMock([mock_pdf])
    res = await _resolve_queryset_candidate(q, 'replace_with')
    print('Resolved:', res)

if __name__ == '__main__':
    asyncio.run(main())
