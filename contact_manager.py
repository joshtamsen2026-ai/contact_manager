"""Manage contacts with SQLAlchemy and demonstrate CRUD operations.

Run ``python contact_manager.py`` to create and open CSV and Word reports.
Use ``--no-open`` to generate the reports without opening desktop applications.

The six functions use a persistent ``contacts.db`` by default. The demonstration
uses a separate in-memory database, so rerunning it cannot change saved contacts.

Reading guide:
    1. Imports and paths select the libraries and files used by the program.
    2. Contact describes the database table; the CRUD functions work with it.
    3. run_demo records sample operations without changing saved contacts.
    4. Report functions export the recorded results; main ties everything together.

Triple-quoted strings like this one are docstrings, which describe modules,
classes, and functions. Lines beginning with ``##`` explain individual steps.
"""

## Postpone evaluation of type hints so they can refer to classes defined later.
from __future__ import annotations

## Python's standard library handles command-line options, files, and launching
## reports. A dataclass stores related values; Path builds paths safely.
import argparse
import csv
import io
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

## SQLAlchemy maps Python objects to SQLite rows and builds database queries.
from sqlalchemy import Boolean, String, create_engine, select
from sqlalchemy.engine import Engine
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column

## Put outputs beside this script, regardless of the current working directory.
BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / "contact_manager_results.csv"
DOCX_PATH = BASE_DIR / "contact_manager_report.docx"
## The default database is persistent; create_engine does not open it yet.
engine = create_engine(f"sqlite:///{(BASE_DIR / 'contacts.db').as_posix()}")


## Database model: SQLite enforces the required and unique column constraints.
class Base(DeclarativeBase):
    """Base class shared by SQLAlchemy models."""


class Contact(Base):
    """Describe one contact and its columns in the ``contacts`` table.

    Each instance represents one row. The ID is its primary key; email is
    required and unique, while a missing phone is stored as SQL NULL.
    """

    ## SQLAlchemy uses this name when creating and querying the SQLite table.
    __tablename__ = "contacts"

    ## Mapped describes the Python type; mapped_column configures the SQL column.
    id: Mapped[int] = mapped_column(primary_key=True)
    first_name: Mapped[str] = mapped_column(String, nullable=False)
    last_name: Mapped[str] = mapped_column(String, nullable=False)
    email: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    phone: Mapped[str | None] = mapped_column(String, nullable=True)
    ## The ORM supplies False on insertion when no favorite value is provided.
    favorite: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)


def _session(database: Engine) -> Session:
    """Create the table if needed and return a session for one operation.

    Args:
        database: The SQLAlchemy engine connected to the chosen database.

    Returns:
        A session that callers should close with a ``with`` block.
    """
    ## create_all leaves existing tables and rows intact.
    Base.metadata.create_all(database)
    ## Keep loaded values available after a commit and the session closes.
    return Session(database, expire_on_commit=False)


def add_contact(
    first_name: str,
    last_name: str,
    email: str,
    phone: str | None = None,
    *,
    database: Engine = engine,
) -> Contact:
    """Add and return a contact, committing it to the selected database.

    Args:
        first_name: Required first name; surrounding whitespace is removed.
        last_name: Required last name; surrounding whitespace is removed.
        email: Required unique email; surrounding whitespace is removed.
        phone: Optional phone number, or None when not provided.
        database: Optional keyword-only engine; defaults to the persistent file.

    Raises:
        ValueError: A required field is blank.
        sqlalchemy.exc.IntegrityError: The email already exists.
    """
    ## Reject blank required fields before SQLite receives the new row.
    if not first_name.strip() or not last_name.strip() or not email.strip():
        raise ValueError("First name, last name, and email must not be blank.")
    ## Build a mapped object. The database assigns its ID during the commit.
    contact = Contact(
        first_name=first_name.strip(),
        last_name=last_name.strip(),
        email=email.strip(),
        phone=phone,
    )
    ## The context manager closes the session; commit saves the new row.
    with _session(database) as session:
        session.add(contact)
        session.commit()
    return contact


def list_contacts(*, database: Engine = engine) -> list[Contact]:
    """Return every contact, sorted by last name, first name, and ID.

    Args:
        database: Optional keyword-only engine; defaults to the persistent file.
    """
    with _session(database) as session:
        ## Secondary sort keys keep equal names in a predictable order.
        statement = select(Contact).order_by(
            Contact.last_name, Contact.first_name, Contact.id
        )
        return list(session.scalars(statement))


def find_contact(email: str, *, database: Engine = engine) -> Contact | None:
    """Find one contact by exact email, or return None when it is absent.

    Args:
        email: The email to look up.
        database: Optional keyword-only engine; defaults to the persistent file.
    """
    with _session(database) as session:
        ## scalar returns the first matching object, or None for no match.
        return session.scalar(select(Contact).where(Contact.email == email))


def update_phone(
    email: str, new_phone: str | None, *, database: Engine = engine
) -> Contact | None:
    """Update a contact's phone and return it, or None if not found.

    Args:
        email: Exact email of the contact to update.
        new_phone: Replacement phone number; None clears the phone.
        database: Optional keyword-only engine; defaults to the persistent file.
    """
    with _session(database) as session:
        ## Change and commit only when the email identifies a saved contact.
        contact = session.scalar(select(Contact).where(Contact.email == email))
        if contact is not None:
            contact.phone = new_phone
            session.commit()
        return contact


def toggle_favorite(email: str, *, database: Engine = engine) -> Contact | None:
    """Flip and return a contact's favorite status, or return None if absent.

    Args:
        email: Exact email of the contact to update.
        database: Optional keyword-only engine; defaults to the persistent file.
    """
    with _session(database) as session:
        contact = session.scalar(select(Contact).where(Contact.email == email))
        if contact is not None:
            ## Python's not converts True to False and False to True.
            contact.favorite = not contact.favorite
            session.commit()
        return contact


def delete_contact(email: str, *, database: Engine = engine) -> bool:
    """Delete a contact by email and report whether it existed.

    Args:
        email: Exact email of the contact to remove.
        database: Optional keyword-only engine; defaults to the persistent file.

    Returns:
        True after a successful deletion, or False if no contact matched.
    """
    with _session(database) as session:
        contact = session.scalar(select(Contact).where(Contact.email == email))
        ## Do not issue a deletion if the query found no row.
        if contact is None:
            return False
        ## A deletion is not permanent until the session commits it.
        session.delete(contact)
        session.commit()
        return True


## Immutable snapshots keep the initial list unchanged after later updates.
@dataclass(frozen=True)
class ContactSnapshot:
    """Keep an immutable copy of a contact at one point in the demonstration.

    Unlike a database object, a snapshot does not change when a later operation
    edits or deletes the corresponding row. Its fields mirror the Contact model.
    """

    id: int
    first_name: str
    last_name: str
    email: str
    phone: str | None
    favorite: bool

    @classmethod
    def from_contact(cls, contact: Contact) -> ContactSnapshot:
        """Copy the visible fields from a database contact into a snapshot."""
        ## cls constructs a new snapshot without retaining the database session.
        return cls(
            contact.id,
            contact.first_name,
            contact.last_name,
            contact.email,
            contact.phone,
            contact.favorite,
        )


@dataclass(frozen=True)
class DemoResults:
    """Hold the initial and final lists and descriptions of each operation.

    Tuples and frozen dataclasses prevent accidental changes while reporting.
    """

    initial: tuple[ContactSnapshot, ...]
    final: tuple[ContactSnapshot, ...]
    actions: tuple[tuple[str, str], ...]


def run_demo() -> DemoResults:
    """Demonstrate all six CRUD operations in an isolated temporary database.

    Returns:
        Before-and-after snapshots and plain-English operation results.

    Raises:
        RuntimeError: An expected find, update, toggle, or delete did not work.
    """
    ## An in-memory SQLite database disappears when this engine is disposed;
    ## the demonstration therefore cannot delete contacts from contacts.db.
    demo_engine = create_engine("sqlite:///:memory:")
    try:
        ## Each tuple contains first name, last name, email, and optional phone.
        samples = (
            ("Ava", "Chen", "ava@example.com", "555-0101"),
            ("Noah", "Patel", "noah@example.com", None),
            ("Mia", "Rivera", "mia@example.com", "555-0103"),
            ("Liam", "Brooks", "liam@example.com", "555-0104"),
            ("Zoe", "Taylor", "zoe@example.com", None),
            ## Repeated names use distinct email addresses to satisfy uniqueness.
            ("Ava", "Chen", "ava.chen2@example.com", "555-0111"),
            ("Mia", "Rivera", "mia.rivera2@example.com", "555-0112"),
            ("Liam", "Brooks", "liam.brooks2@example.com", "555-0113"),
            ("Elena", "Cruz", "elena.cruz@example.com", "555-0114"),
            ("Owen", "Foster", "owen.foster@example.com", "555-0115"),
            ("Nora", "Garcia", "nora.garcia@example.com", "555-0116"),
            ("Ethan", "Hughes", "ethan.hughes@example.com", "555-0117"),
            ("Layla", "James", "layla.james@example.com", "555-0118"),
            ("Lucas", "Kim", "lucas.kim@example.com", "555-0119"),
            ("Aria", "Morgan", "aria.morgan@example.com", "555-0120"),
            ("Mason", "Nguyen", "mason.nguyen@example.com", "555-0121"),
            ("Sofia", "Owens", "sofia.owens@example.com", "555-0122"),
            ("Leo", "Perez", "leo.perez@example.com", "555-0123"),
            ("Chloe", "Reed", "chloe.reed@example.com", "555-0124"),
            ("Isaac", "Stone", "isaac.stone@example.com", "555-0125"),
        )
        ## Add each example using the same public function used by real callers.
        for first, last, email, phone in samples:
            add_contact(first, last, email, phone, database=demo_engine)
        ## Save a separate copy of the sorted list before making changes.
        initial = tuple(
            ContactSnapshot.from_contact(c) for c in list_contacts(database=demo_engine)
        )
        ## Exercise lookup, phone update, favorite toggle, and deletion by email.
        found = find_contact("ava@example.com", database=demo_engine)
        updated = update_phone("noah@example.com", "555-0199", database=demo_engine)
        favorite = toggle_favorite("mia@example.com", database=demo_engine)
        deleted = delete_contact("zoe@example.com", database=demo_engine)
        ## Stop rather than generate a misleading report if a step failed.
        if found is None or updated is None or favorite is None or not deleted:
            raise RuntimeError(
                "The contact demonstration did not complete as expected."
            )
        ## Capture the sorted list again after the updates and deletion.
        final = tuple(
            ContactSnapshot.from_contact(c) for c in list_contacts(database=demo_engine)
        )
        ## These descriptions supply the console, CSV, and Word report sections.
        actions = (
            (
                "Add",
                (
                    f"Added {len(initial)} sample contacts; three names repeat with "
                    "different emails. Favorites default to False."
                ),
            ),
            ("List before changes", "Listed all contacts in last-name order."),
            ("Find", f"Found {found.first_name} {found.last_name} by {found.email}."),
            ("Update", f"Changed {updated.email}'s phone to {updated.phone}."),
            (
                "Toggle favorite",
                f"Set {favorite.email}'s favorite status to {favorite.favorite}.",
            ),
            ("Delete", "Deleted zoe@example.com."),
            (
                "List after changes",
                f"Listed {len(final)} remaining contacts in last-name order.",
            ),
        )
        return DemoResults(initial, final, actions)
    finally:
        ## Always release the demo database, including when a step raises an error.
        demo_engine.dispose()


def _summary(results: DemoResults) -> str:
    """Explain the measured results, model requirements, and demo scope.

    Args:
        results: Snapshots and actions collected by run_demo.

    Returns:
        A paragraph shared by the console and both reports.
    """
    ## Boolean expressions count as 1 for True and 0 for False in sum().
    favorites = sum(contact.favorite for contact in results.final)
    with_phone = sum(contact.phone is not None for contact in results.final)
    return (
        f"The isolated demonstration added {len(results.initial)} contacts and "
        f"finished with {len(results.final)} after one deletion. It found a contact "
        f"by email, updated one phone number, and toggled one favorite; the final "
        f"list contains {favorites} favorite(s) and {with_phone} contact(s) with "
        "a phone number. Both lists are sorted by last name. The model requires "
        "first name, last name, and a unique email; phone is optional and favorite "
        "defaults to False. The demonstration does not modify the persistent "
        "contacts.db used by the CRUD functions."
    )


def _contact_values(contact: ContactSnapshot) -> tuple[str, ...]:
    """Format one snapshot as text for the console, CSV, and Word tables."""
    ## Convert the ID and boolean to readable text; explain a missing phone.
    return (
        str(contact.id),
        contact.first_name,
        contact.last_name,
        contact.email,
        contact.phone if contact.phone is not None else "(not provided)",
        "Yes" if contact.favorite else "No",
    )


## Shared headers keep all three outputs in the same column order.
HEADERS = ("ID", "First name", "Last name", "Email", "Phone", "Favorite")
INITIAL_CONTACTS_LABEL = "Initial contacts"
FINAL_CONTACTS_LABEL = "Final contacts"


def _spreadsheet_safe(value: str) -> str:
    """Prefix formula-like text so spreadsheet software treats it as text."""
    ## A leading apostrophe prevents common spreadsheet formula prefixes from
    ## being interpreted as formulas when the CSV is opened in an application.
    stripped = value.lstrip()
    leading_whitespace = value[: len(value) - len(stripped)]
    if any(char in leading_whitespace for char in "\t\r\n") or stripped.startswith(
        ("=", "+", "-", "@")
    ):
        return "'" + value
    return value


def write_csv(path: Path, results: DemoResults) -> None:
    """Write an overview, operations, contact lists, and summary to a CSV file.

    Args:
        path: Destination file; an existing file is replaced.
        results: Demonstration data to include in the report.
    """
    ## utf-8-sig helps spreadsheet applications recognize UTF-8; newline=""
    ## lets csv.writer handle line endings without adding blank rows on Windows.
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        ## Keep descriptions near their labels while preserving a single header
        ## and consistent columns for spreadsheet filtering and CSV readers.
        writer.writerow(("Section", "Item", "Explanation", *HEADERS))

        def write_row(
            section: str, item: str, explanation: str, values: tuple[str, ...] = ()
        ) -> None:
            """Write one spreadsheet-safe row, filling unused contact columns."""
            row = (
                section,
                item,
                explanation,
                *values,
                *("" for _ in HEADERS[len(values) :]),
            )
            writer.writerow(tuple(_spreadsheet_safe(value) for value in row))

        ## Put key figures first so readers can understand the outcome quickly.
        overview = (
            (INITIAL_CONTACTS_LABEL, str(len(results.initial))),
            (FINAL_CONTACTS_LABEL, str(len(results.final))),
            (
                "Final favorites",
                str(sum(contact.favorite for contact in results.final)),
            ),
            (
                "Final contacts with phone",
                str(sum(contact.phone is not None for contact in results.final)),
            ),
        )
        for item, value in overview:
            write_row("Overview", item, value)
        ## Summarize each operation before listing the affected contacts.
        for action, explanation in results.actions:
            write_row("Operations", action, explanation)
        ## Show the final list first, then the original list for comparison.
        for section, contacts in (
            ("Final list", results.final),
            ("Initial list", results.initial),
        ):
            for contact in contacts:
                write_row(section, "Contact", "", _contact_values(contact))
        ## Finish with the explanation of findings and requirements.
        write_row("Summary", "Findings and requirements", _summary(results))


def _chart(results: DemoResults) -> io.BytesIO:
    """Create an in-memory PNG chart comparing initial and final counts.

    Args:
        results: Demonstration data containing both contact lists.

    Returns:
        A memory-backed file positioned at the start of the PNG image.
    """
    ## Import chart libraries only when a Word report needs a chart.
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.ticker import MaxNLocator

    ## The Agg canvas renders without opening a graphical chart window.
    figure = Figure(figsize=(7, 3.5))
    FigureCanvasAgg(figure)
    axes = figure.add_subplot()
    ## Compare the snapshot sizes and label each bar with its count.
    counts = [len(results.initial), len(results.final)]
    bars = axes.bar(
        ["Before changes", "After changes"], counts, color=["#3274a1", "#3a8b67"]
    )
    axes.bar_label(bars)
    axes.set(title="Contact count before and after deletion", ylabel="Contacts")
    ## Contact counts are whole numbers; leave space above the tallest label.
    axes.yaxis.set_major_locator(MaxNLocator(integer=True))
    axes.set_ylim(0, max(counts, default=0) + 1)
    figure.tight_layout()
    ## Store the PNG in memory instead of creating a separate temporary file.
    image = io.BytesIO()
    ## Lower PNG compression reduces rendering time without losing image quality.
    figure.savefig(image, format="png", dpi=150, pil_kwargs={"compress_level": 1})
    ## Rewind so the Word library can read the complete image from the start.
    image.seek(0)
    return image


def write_docx(path: Path, results: DemoResults) -> None:
    """Write a Word report with operations, tables, a chart, and a summary.

    Args:
        path: Destination file; an existing file is replaced.
        results: Demonstration data to include in the report.
    """
    ## Import the Word library only when creating the report.
    from docx import Document
    from docx.shared import Inches, Pt

    ## Start a document with a title, introductory text, and readable font size.
    document = Document()
    document.styles["Normal"].font.size = Pt(10)
    document.add_heading("Contact Manager Report", 0)
    document.add_paragraph(
        "CRUD demonstration using fictional contacts in an isolated in-memory database."
    )
    ## Describe the model requirements before presenting the operation results.
    document.add_heading("Requirements and model", level=1)
    document.add_paragraph(
        "Contact has an ID, required first and last names, and a required unique "
        "email. Phone is optional; favorite is a boolean that defaults to False. "
        "The demonstration exercises add, list, find, update, toggle, and delete."
    )
    ## Give each demonstration action a heading and explanation.
    document.add_heading("Operations and results", level=1)
    for action, explanation in results.actions:
        document.add_heading(action, level=2)
        document.add_paragraph(explanation)
    ## Build separate before-and-after tables with the same shared headers.
    for title, contacts in (
        (INITIAL_CONTACTS_LABEL, results.initial),
        (FINAL_CONTACTS_LABEL, results.final),
    ):
        document.add_heading(title, level=1)
        table = document.add_table(rows=1, cols=len(HEADERS))
        table.style = "Table Grid"
        ## zip pairs each cell with its value; strict=True checks equal lengths.
        for cell, header in zip(table.rows[0].cells, HEADERS, strict=True):
            cell.text = header
        for contact in contacts:
            for cell, value in zip(
                table.add_row().cells, _contact_values(contact), strict=True
            ):
                cell.text = value
    ## Add the chart and explain what the difference in counts means.
    document.add_heading("Contact count", level=1)
    document.add_picture(_chart(results), width=Inches(6))
    document.add_paragraph(
        "The decrease of one contact reflects the demonstrated deletion."
    )
    ## End with the same findings paragraph used in the CSV and console.
    document.add_heading("Summary", level=1)
    document.add_paragraph(_summary(results))
    document.save(str(path))


def open_file(path: Path) -> None:
    """Ask the operating system to open a report in its default application.

    Args:
        path: Existing report file to open.
    """
    ## Each operating system provides a different way to open a document.
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def main() -> int:
    """Run the demonstration, print results, and generate both reports.

    Returns:
        Exit code 0 on success, or 1 if writing or opening a report fails.
        Other errors are allowed to surface instead of being reported as success.
    """
    ## Read --no-open so callers can create reports without launching apps.
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--no-open", action="store_true", help="Do not open reports.")
    args = parser.parse_args()
    ## Run the isolated demonstration once; use its results for every output.
    results = run_demo()
    ## Print each action and its explanation before showing the two snapshots.
    for action, explanation in results.actions:
        print(f"{action}: {explanation}")
    ## Show both contact lists in the same column order as the reports.
    for title, contacts in (
        (INITIAL_CONTACTS_LABEL, results.initial),
        (FINAL_CONTACTS_LABEL, results.final),
    ):
        print(f"\n{title} (sorted by last name):")
        print(" | ".join(HEADERS))
        for contact in contacts:
            print(" | ".join(_contact_values(contact)))
    ## Print the findings, then save the CSV and Word versions beside the script.
    print(f"\n{_summary(results)}")
    ## Surface file errors with the affected path, such as a report locked by
    ## a spreadsheet or Word application. Earlier reports may already be saved.
    for path, writer in ((CSV_PATH, write_csv), (DOCX_PATH, write_docx)):
        try:
            writer(path, results)
        except OSError as error:
            print(f"Could not write {path}: {error}", file=sys.stderr)
            return 1
    print(f"\nCSV report: {CSV_PATH}\nWord report: {DOCX_PATH}")
    ## Unless requested otherwise, ask the OS to open each completed report.
    if not args.no_open:
        for path in (CSV_PATH, DOCX_PATH):
            try:
                open_file(path)
            except OSError as error:
                ## Report opening errors explicitly instead of claiming success.
                print(f"Could not open {path}: {error}", file=sys.stderr)
                return 1
    return 0


## Only run the demonstration when this file is executed directly. Importing
## it from another Python file leaves callers free to use the CRUD functions.
if __name__ == "__main__":
    raise SystemExit(main())
