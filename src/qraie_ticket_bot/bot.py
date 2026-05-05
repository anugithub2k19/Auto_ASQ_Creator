from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from playwright.sync_api import Browser, Locator, Page, Playwright, sync_playwright

from qraie_ticket_bot.config import Config
from qraie_ticket_bot.excel_io import TicketRow


@dataclass(frozen=True)
class CreateResult:
    status: str
    error: str = ""


class QRaieBot:
    def __init__(self, config: Config):
        self.config = config

    def _first_visible(self, locators: list[Locator]) -> Locator:
        for loc in locators:
            try:
                if loc.first.is_visible():
                    return loc.first
            except Exception:  # noqa: BLE001
                continue
        return locators[0].first

    def _modal(self, page: Page) -> Locator:
        return page.locator(".ant-modal-content, .modal-content, [role='dialog']").first

    def _click_field_anchor(self, page: Page, label: str) -> Locator:
        modal = self._modal(page)
        # Find the label, then get its parent row/container, then find the input/select within it.
        lbl = modal.get_by_text(label, exact=False).first
        if lbl.is_visible():
            # Go up to the form row/container, then find the control.
            row = lbl.locator("xpath=ancestor::*[contains(@class,'row') or contains(@class,'form-item') or self::div][1]")
            if row.count() > 0:
                ctrl = row.locator("input, [role='combobox'], .ant-select, textarea").first
                if ctrl.count() > 0:
                    return ctrl
        # Fallback: find any input near the label text.
        return modal.locator(f"label:has-text('{label}') + * input, label:has-text('{label}') + * [role='combobox'], label:has-text('{label}') + * .ant-select").first

    def _select_dropdown_value(self, page: Page, *, label: str, value: str) -> None:
        value = (value or "").strip()
        if not value:
            raise ValueError(f"Excel value missing for '{label}'")

        # 1) Try get_by_role combobox with label name.
        try:
            combobox = page.get_by_role("combobox", name=label, exact=False)
            if combobox.count() > 0 and combobox.first.is_visible():
                combobox.first.click()
                page.wait_for_timeout(800)
                search = combobox.first.locator("..").locator(".ant-select-selection-search-input, input[role='combobox']").first
                if search.is_visible():
                    search.fill(value)
                    page.wait_for_timeout(800)
                opt = page.locator(".ant-select-item-option").filter(has_text=value).first
                if opt.is_visible(timeout=2000):
                    opt.click()
                    page.wait_for_timeout(1000)
                    return
        except Exception:  # noqa: BLE001
            pass

        # 2) Try native select by label.
        try:
            native = page.get_by_label(label, exact=False)
            if native.count() > 0 and native.first.is_visible():
                native.first.select_option(label=value)
                page.wait_for_timeout(500)
                return
        except Exception:  # noqa: BLE001
            pass

        # 3) Find the ant-select by locating the label element.
        label_elem = page.locator(f"label:has-text('{label}')").first
        if label_elem.count() > 0:
            try:
                for_attr = label_elem.get_attribute("for", timeout=1000)
                if for_attr:
                    select_box = page.locator(f"#{for_attr}, [id='{for_attr}'] >> .ant-select").first
                    if select_box.count() == 0:
                        select_box = page.locator(f"#{for_attr}").first
                    try:
                        current_val = select_box.locator(".ant-select-selection-item").first.inner_text(timeout=1000)
                        if value.lower() in current_val.lower():
                            return
                    except Exception:  # noqa: BLE001
                        pass
                    select_box.click()
                    page.wait_for_timeout(800)
                    # Try to type in search if available
                    try:
                        search = select_box.locator(".ant-select-selection-search-input, input[role='combobox']").first
                        if search.is_visible():
                            search.fill(value)
                            page.wait_for_timeout(800)
                    except Exception:  # noqa: BLE001
                        pass
                    opt = page.locator(".ant-select-item-option").filter(has_text=value).first
                    if opt.is_visible(timeout=2000):
                        opt.click()
                        page.wait_for_timeout(1000)
                        return
            except Exception:  # noqa: BLE001
                pass

            row = label_elem.locator("xpath=ancestor::*[contains(@class, 'ant-row')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[contains(@class, 'ant-form-item')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[2]")
            select_box = row.locator(".ant-select").first
        else:
            label_loc = page.locator(f"text={label}").first
            row = label_loc.locator("xpath=ancestor::div[2]")
            select_box = row.locator(".ant-select").first

        try:
            current_val = select_box.locator(".ant-select-selection-item").first.inner_text(timeout=2000)
            if value.lower() in current_val.lower():
                return
        except Exception:  # noqa: BLE001
            pass

        select_box.click()
        page.wait_for_timeout(800)

        # Try to type in search if available (some dropdowns don't have search)
        try:
            search = select_box.locator(".ant-select-selection-search-input, input[role='combobox']").first
            if search.is_visible():
                search.fill(value)
                page.wait_for_timeout(800)
        except Exception:  # noqa: BLE001
            pass

        # Look for the option and click it
        opt = page.locator(".ant-select-item-option").filter(has_text=value).first
        if opt.count() > 0:
            try:
                if opt.is_visible(timeout=2000):
                    opt.click()
                    page.wait_for_timeout(1000)
                    return
            except Exception:  # noqa: BLE001
                pass

        # If no option found, close dropdown gracefully
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)

    def _fill_text(self, page: Page, *, label: str, value: str) -> None:
        value = (value or "").strip()
        if not value:
            raise ValueError(f"Excel value missing for '{label}'")

        # 1) Try get_by_label first.
        try:
            input_field = page.get_by_label(label, exact=False)
            if input_field.count() > 0 and input_field.first.is_visible():
                input_field.first.click()
                page.wait_for_timeout(200)
                input_field.first.fill(value)
                page.wait_for_timeout(300)
                return
        except Exception:  # noqa: BLE001
            pass

        # 2) Find the input by locating the label element and its row.
        label_elem = page.locator(f"label:has-text('{label}')").first
        if label_elem.count() > 0:
            row = label_elem.locator("xpath=ancestor::*[contains(@class, 'ant-row')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[contains(@class, 'ant-form-item')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[2]")
            input_field = row.locator("input:not([type='search']):not([role='combobox'])").first
        else:
            label_loc = page.locator(f"text={label}").first
            row = label_loc.locator("xpath=ancestor::div[2]")
            input_field = row.locator("input:not([type='search']):not([role='combobox'])").first

        try:
            input_field.click()
            page.wait_for_timeout(200)
            input_field.fill(value)
            page.wait_for_timeout(300)
        except Exception:  # noqa: BLE001
            page.locator(f"input[placeholder*='{label}' i]").first.fill(value)
            page.wait_for_timeout(300)

    def _fill_rich_text_description(self, page: Page, value: str) -> None:
        value = (value or "").strip()
        if not value:
            raise ValueError("Excel value missing for 'Description'")

        label_elem = page.locator("label:has-text('Description')").first
        if label_elem.count() > 0:
            row = label_elem.locator("xpath=ancestor::*[contains(@class, 'ant-row')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[contains(@class, 'ant-form-item')][1]")
            if row.count() == 0:
                row = label_elem.locator("xpath=ancestor::div[2]")
        else:
            label_loc = page.locator("text=Description").first
            row = label_loc.locator("xpath=ancestor::div[2]")

        candidates = [
            row.locator("[contenteditable='true']").first,
            row.locator("textarea").first,
            page.locator("[contenteditable='true']").first,
        ]

        target = self._first_visible(candidates)
        target.click()
        page.wait_for_timeout(300)
        try:
            target.fill(value)
        except Exception:  # noqa: BLE001
            page.keyboard.press("Control+A")
            page.keyboard.type(value)
        page.wait_for_timeout(300)

    def _fill_login_username(self, page: Page, value: str) -> None:
        self._first_visible(
            [
                # Preferred: placeholder or explicit name
                page.get_by_placeholder("Username", exact=False),
                page.locator('input[name="username"]'),
                # If the UI shows "Username" as plain text, grab the nearest following input
                page.locator("text=Username").locator("xpath=following::input[1]"),
                # Accessibility label (works only if properly wired)
                page.get_by_label("Username", exact=False),
                # Fallbacks
                page.locator('input[type="email"]'),
                page.locator('input[type="text"]'),
            ]
        ).fill(value)

    def _fill_login_password(self, page: Page, value: str) -> None:
        self._first_visible(
            [
                page.locator('input[type="password"]'),
                page.get_by_placeholder("Password", exact=False),
                page.locator('input[name="password"]'),
                page.locator("text=Password").locator("xpath=following::input[1]"),
                page.get_by_label("Password", exact=False),
            ]
        ).fill(value)

    def run(self, tickets: list[TicketRow], progress_callback=None) -> dict[int, CreateResult]:
        results: dict[int, CreateResult] = {}
        artifacts_dir = Path(self.config.output.artifacts_dir)
        artifacts_dir.mkdir(parents=True, exist_ok=True)

        with sync_playwright() as p:
            browser = self._launch(p)
            context = self._new_context(browser)
            page = context.new_page()
            page.set_default_timeout(self.config.run.timeout_ms)
            context.tracing.start(screenshots=True, snapshots=True, sources=False)

            try:
                try:
                    self._login(page)
                except Exception as e:  # noqa: BLE001
                    try:
                        page.screenshot(path=str(artifacts_dir / "startup_error.png"), full_page=True)
                    except Exception:  # noqa: BLE001
                        pass
                    try:
                        context.tracing.stop(path=str(artifacts_dir / "startup_trace.zip"))
                    except Exception:  # noqa: BLE001
                        pass
                    raise RuntimeError(
                        "Startup failed (login/navigation). "
                        "Check output/artifacts/startup_error.png and startup_trace.zip"
                    ) from e

                total = len(tickets)
                for i, t in enumerate(tickets):
                    if not t.run:
                        results[t.row_index_1based] = CreateResult("SKIPPED", "")
                        continue
                    try:
                        self._create_ticket(page, t)
                        results[t.row_index_1based] = CreateResult("CREATED", "")
                    except Exception as e:  # noqa: BLE001
                        shot = artifacts_dir / f"row_{t.row_index_1based}.png"
                        try:
                            page.screenshot(path=str(shot), full_page=True)
                        except Exception:  # noqa: BLE001
                            pass
                        results[t.row_index_1based] = CreateResult("FAILED", str(e))
                    
                    if progress_callback:
                        progress_callback(i + 1, total, results[t.row_index_1based].status)
            finally:
                try:
                    context.tracing.stop(path=str(artifacts_dir / "run_trace.zip"))
                except Exception:  # noqa: BLE001
                    pass
                context.close()
                browser.close()

        return results

    def _launch(self, p: Playwright) -> Browser:
        return p.chromium.launch(
            headless=self.config.run.headless,
            slow_mo=self.config.run.slow_mo_ms,
        )

    def _new_context(self, browser: Browser):
        context_kwargs = {}
        if self.config.run.grant_geolocation:
            context_kwargs["geolocation"] = {
                "latitude": self.config.run.geolocation_lat,
                "longitude": self.config.run.geolocation_lon,
            }
            context_kwargs["permissions"] = ["geolocation"]

        if self.config.auth.mode == "storage_state":
            state_path = Path(self.config.auth.storage_state_path)
            if not state_path.exists():
                raise FileNotFoundError(
                    f"storage_state_path not found: {state_path}. "
                    f"Either set auth.mode=credentials, or generate storage state."
                )
            return browser.new_context(storage_state=str(state_path), **context_kwargs)
        return browser.new_context(**context_kwargs)

    def _login(self, page: Page) -> None:
        page.goto(self.config.app.url, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        page.wait_for_url("**/workplace/**", timeout=15_000)

        # If the app displays a location-permission modal, accept it.
        try:
            page.get_by_role("button", name="Allow Location Access", exact=False).click(timeout=3_000)
        except Exception:  # noqa: BLE001
            pass

        if self.config.auth.mode == "storage_state":
            return

        if not self.config.auth.username or not self.config.auth.password:
            raise ValueError("Missing auth.username/auth.password in config.yaml")

        # Based on the provided login screenshot: fields labeled "Username" and "Password",
        # and a button with text "Login".
        self._fill_login_username(page, self.config.auth.username)
        self._fill_login_password(page, self.config.auth.password)
        page.get_by_role("button", name="Login", exact=False).click()

        # Post-login: avoid waiting on sidebar labels (can be hidden on desktop).
        # Instead, wait for the login button to disappear OR for any authenticated UI anchor.
        page.wait_for_load_state("domcontentloaded")
        try:
            page.get_by_role("button", name="Login", exact=False).wait_for(state="hidden", timeout=20_000)
            return
        except Exception:  # noqa: BLE001
            pass

        # Fallback: "Add New" is visible on the Tasks/ASQ list screen.
        try:
            page.get_by_role("button", name="Add New", exact=False).first.wait_for(timeout=20_000)
            return
        except Exception as e:  # noqa: BLE001
            raise RuntimeError("Login did not reach the authenticated area.") from e

    def _open_create_ticket_form(self, page: Page) -> None:
        # Ensure we're on the right page.
        try:
            page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:  # noqa: BLE001
            pass

        # Wait for any existing modal to close completely.
        try:
            page.locator(".ant-modal-wrap").first.wait_for(state="hidden", timeout=5_000)
        except Exception:  # noqa: BLE001
            pass

        # Force close any stubborn modal.
        for _ in range(3):
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(300)
            except Exception:  # noqa: BLE001
                pass

        # Try clicking modal close button if visible.
        try:
            close_btn = page.locator(".ant-modal-close, .ant-modal-close-x, button:has(.anticon-close), [aria-label='Close']").first
            if close_btn.is_visible(timeout=2000):
                close_btn.click()
                page.wait_for_timeout(500)
        except Exception:  # noqa: BLE001
            pass

        # Wait for modal to be hidden.
        try:
            page.locator(".ant-modal-wrap").first.wait_for(state="hidden", timeout=3_000)
        except Exception:  # noqa: BLE001
            pass

        # Click Add New to open the form.
        add_new = page.get_by_role("button", name="Add New", exact=False).first
        add_new.wait_for(state="visible", timeout=15_000)
        add_new.click()

        # Wait for the form modal to appear.
        page.locator(".ant-modal-content, [role='dialog']").get_by_text("Create New Asq", exact=False).first.wait_for(timeout=15_000)
        page.wait_for_timeout(1000)
        
        # Explicitly set focus to Tenant dropdown first
        try:
            tenant_select = page.locator(".ant-modal-content .ant-select").first
            tenant_select.click()
            page.wait_for_timeout(500)
            page.keyboard.press("Escape")  # Close dropdown without selecting
            page.wait_for_timeout(300)
        except Exception:  # noqa: BLE001
            pass

    def _create_ticket(self, page: Page, t: TicketRow) -> None:
        # Ensure page is still valid
        if not page.context:
            raise RuntimeError("Browser context was closed")
            
        self._open_create_ticket_form(page)
        page.wait_for_timeout(1000)

        # Step 1: Tenant dropdown (typable) - focus is already here when modal opens
        self._fill_typable_dropdown(page, value=t.tenant)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 2: Project dropdown (typable)
        self._fill_typable_dropdown(page, value=t.project)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 3: Title (free text)
        self._fill_text_field(page, value=t.title)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 4: Module dropdown (typable)
        self._fill_typable_dropdown(page, value=t.module)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 5: Severity (non-typable, keep default "Low")
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 6: Priority dropdown (typable)
        self._fill_typable_dropdown(page, value=t.priority)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 7: Issue Category (non-typable dropdown)
        self._fill_non_typable_dropdown(page, value=t.issue_category)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 8: Sub Category dropdown (typable)
        self._fill_typable_dropdown(page, value=t.sub_category)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 9: Owner dropdown (typable)
        self._fill_typable_dropdown(page, value=t.owner)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)

        # Step 10: Description (free text editor)
        # Navigate to Description field
        for _ in range(2):
            page.keyboard.press("Tab")
            page.wait_for_timeout(200)
        
        # Find and fill Description
        try:
            desc_field = page.locator("[contenteditable='true']").first
            desc_field.click()
            page.wait_for_timeout(300)
            page.keyboard.press("Control+A")
            page.wait_for_timeout(200)
            page.keyboard.type(t.description)
        except Exception:  # noqa: BLE001
            pass
        page.wait_for_timeout(800)

        # Step 11: Click Create Ticket button
        try:
            create_btn = page.get_by_role("button", name="Create Ticket", exact=False)
            if create_btn.is_visible(timeout=5000):
                create_btn.click()
                page.wait_for_timeout(2000)
                try:
                    page.get_by_text("success", exact=False).first.wait_for(timeout=5_000)
                except Exception:  # noqa: BLE001
                    pass
        except Exception:  # noqa: BLE001
            pass
        
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        # Close modal completely - use multiple strategies
        # Strategy 1: Click Cancel button if visible (sometimes modal stays open after success)
        try:
            cancel_btn = page.get_by_role("button", name="Cancel", exact=False).first
            if cancel_btn.is_visible(timeout=2000):
                cancel_btn.click()
                page.wait_for_timeout(1000)
        except Exception:  # noqa: BLE001
            pass

        # Strategy 2: Click close button (X)
        try:
            close_btn = page.locator(".ant-modal-close, [aria-label='Close'], button:has(.anticon-close)").first
            if close_btn.is_visible(timeout=2000):
                close_btn.click()
                page.wait_for_timeout(1000)
        except Exception:  # noqa: BLE001
            pass

        # Strategy 3: Press Escape multiple times
        for _ in range(3):
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(300)
            except Exception:  # noqa: BLE001
                pass

        # Strategy 4: Wait for modal to disappear
        try:
            page.locator(".ant-modal-wrap").first.wait_for(state="hidden", timeout=5_000)
        except Exception:  # noqa: BLE001
            pass
            
        page.wait_for_timeout(1000)
        
        # Verify page is still valid after ticket creation
        try:
            page.title()
        except Exception:  # noqa: BLE001
            raise RuntimeError("Page became invalid after ticket creation")

    def _fill_typable_dropdown(self, page: Page, *, value: str) -> None:
        """Fill a dropdown that allows typing (field already has focus from Tab)."""
        value = (value or "").strip()
        if not value:
            return

        # Clear existing value and type new one
        page.keyboard.press("Control+A")
        page.wait_for_timeout(100)
        page.keyboard.type(value)
        page.wait_for_timeout(1500)

        # Click the filtered option
        opt = page.locator(".ant-select-item-option").filter(has_text=value).first
        if opt.count() > 0:
            try:
                if opt.is_visible(timeout=5000):
                    opt.click()
                    page.wait_for_timeout(1000)
                    return
            except Exception:  # noqa: BLE001
                pass

    def _fill_non_typable_dropdown(self, page: Page, *, value: str) -> None:
        """Fill a dropdown that only allows selection (field already has focus from Tab)."""
        value = (value or "").strip()
        if not value:
            return

        # Open dropdown
        page.keyboard.press("Enter")
        page.wait_for_timeout(2000)

        # Find and click the matching option
        opt = page.locator(".ant-select-item-option").filter(has_text=value).first
        if opt.count() > 0:
            try:
                if opt.is_visible(timeout=10000):
                    opt.scroll_into_view_if_needed()
                    page.wait_for_timeout(1000)
                    opt.click()
                    page.wait_for_timeout(1000)
                    return
            except Exception:  # noqa: BLE001
                pass

    def _fill_text_field(self, page: Page, *, value: str) -> None:
        """Fill a regular text input field (field already has focus from Tab)."""
        value = (value or "").strip()
        if not value:
            return

        page.keyboard.press("Control+A")
        page.wait_for_timeout(100)
        page.keyboard.type(value)
        page.wait_for_timeout(300)

