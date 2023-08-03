from datetime import datetime

from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.charity_project import charity_project_crud
from app.crud.donation import donation_crud


async def distribute_for_projects(session: AsyncSession):
    """
    Функция распределения пожертвований по проектам.

    projects - список незакрытых проектов
    project_index - индекс обрабатываемого проекта
    amount_to_complete - сумма, необходимая для закрытия проекта
    donations - список нераспределенных пожертвований
    donation_index - индекс обрабатываемого пожертвования
    unused_amount - неизрасходованный остаток пожертвования
    distribute_amount - распределяемая к конкретной операции сумма
    """
    project_index = 0
    donation_index = 0

    projects = await charity_project_crud.get_unclosed_objects(
        session=session
    )

    donations = await donation_crud.get_unclosed_objects(
        session=session
    )

    while donation_index < len(donations) and project_index < len(projects):
        unused_amount = (donations[donation_index].full_amount -
                         donations[donation_index].invested_amount)

        amount_to_complete = (projects[project_index].full_amount -
                              projects[project_index].invested_amount)

        distribute_amount = min(unused_amount, amount_to_complete)

        donations[donation_index].invested_amount += distribute_amount

        projects[project_index].invested_amount += distribute_amount

        if (donations[donation_index].invested_amount ==
                donations[donation_index].full_amount):
            donations[donation_index].fully_invested = True
            donations[donation_index].close_date = datetime.now()
            donation_index += 1

        if (projects[project_index].invested_amount ==
                projects[project_index].full_amount):
            projects[project_index].fully_invested = True
            projects[project_index].close_date = datetime.now()
            project_index += 1

    session.add_all(donations + projects)
    await session.commit()
