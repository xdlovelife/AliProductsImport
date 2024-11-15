import tkinter as tk
import logging
import time
from tkinter import ttk, simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchWindowException

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 读取保存的Firefox配置文件路径
def read_profile_path():
    try:
        with open("firefox_profile.txt") as file:
            return file.read().strip()
    except FileNotFoundError:
        return ""

# 保存有效路径到文件
def save_profile_path(path):
    with open("firefox_profile.txt", "w") as file:
        file.write(path)

# 保存分类列表到文件
def save_categories(categories):
    with open("categories.txt", "w") as file:
        for category in categories:
            file.write(f"{category}\n")

# 读取分类列表从文件
def read_categories():
    try:
        with open("categories.txt") as file:
            return [line.strip() for line in file.readlines()]
    except FileNotFoundError:
        return []

# 弹窗加载和验证Firefox配置文件路径
def get_valid_profile_path():
    profile_path = read_profile_path()
    if profile_path:
        confirm_reset = messagebox.askyesno("确认配置路径",
                                            f"已设置Firefox配置文件路径为:\n{profile_path}\n\n是否确认使用该路径？\n选择否将重设路径。")
        if confirm_reset:
            return profile_path

    while True:
        path = simpledialog.askstring("Firefox配置文件", "请输入Firefox配置文件路径:")
        if not path:
            messagebox.showerror("错误", "请输入有效的路径。")
        else:
            try:
                # 检查路径是否有效
                options = webdriver.FirefoxOptions()
                options.add_argument(f"-profile {path}")
                driver = webdriver.Firefox(options=options)
                driver.quit()
                save_profile_path(path)
                return path
            except Exception as e:
                messagebox.showerror("错误", f"无效的路径或配置文件: {e}")

# 弹窗加载可供选择的产品分类
def input_product_category(profile_path):
    root = tk.Tk()
    root.title("批量输入产品分类")

    # 设置窗口大小和位置
    window_width = 500
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    categories = read_categories()  # 从文件读取分类列表

    # 更新显示产品分类列表
    def update_category_list():
        listbox.delete(0, tk.END)
        for category in categories:
            listbox.insert(tk.END, category)

    # 添加分类
    def add_category():
        category = category_entry.get().strip()
        if category:
            categories.append(category)
            update_category_list()
            category_entry.delete(0, tk.END)  # 清空输入框内容

    # 删除选定分类
    def delete_category():
        selection = listbox.curselection()
        if selection:
            index = selection[0]
            del categories[index]
            update_category_list()

    # 创建标签和输入框
    ttk.Label(root, text="产品分类列表").pack(pady=10)
    listbox = tk.Listbox(root)
    listbox.pack()

    update_category_list()  # 更新显示分类列表

    # 创建输入框和按钮
    category_entry = ttk.Entry(root)
    category_entry.pack()
    ttk.Button(root, text="添加分类", command=add_category).pack()
    ttk.Button(root, text="删除选定分类", command=delete_category).pack()

    # 定义确认按钮的回调函数
    def confirm_input():
        if categories:
            save_categories(categories)  # 保存分类列表到文件
            root.destroy()  # 销毁主窗口
            open_alibaba(categories, profile_path)  # 将 profile_path 和分类列表作为参数传递
        else:
            messagebox.showwarning("警告", "请输入至少一个分类！")

    # 创建确认按钮
    ttk.Button(root, text="开始执行", command=confirm_input).pack(pady=10)

    root.mainloop()

# 打开阿里巴巴并处理链接
def open_alibaba(selected_categories, profile_path):
    logging.info("打开阿里巴巴页面")
    options = webdriver.FirefoxOptions()
    options.add_argument(f"-profile {profile_path}")  # 设置Firefox配置文件路径
    options.headless = False  # 设置为 False 可以看到浏览器操作过程
    try:
        browser = webdriver.Firefox(options=options)
    except Exception as e:
        logging.error(f"无法启动Firefox: {e}")
        messagebox.showerror("错误", f"无法启动Firefox: {e}")
        return

    success_count = 0  # 初始化成功计数器

    for category in selected_categories:
        try:
            success_count = process_link(browser, "https://www.alibaba.com/", category, success_count)
        except Exception as e:
            logging.error(f"处理分类 {category} 时发生错误: {e}")
            # 如果出现错误，关闭浏览器，重新打开并处理
            browser.quit()
            browser = webdriver.Firefox(options=options)
            success_count = process_link(browser, "https://www.alibaba.com/", category, success_count)

    browser.quit()  # 处理完所有产品后关闭浏览器

def process_link(browser, link, category, success_count):
    logging.info(f"处理分类: {category}")
    try:
        logging.info(f"处理链接: {link}")
        browser.get(link)
        browser.switch_to.window(browser.window_handles[0])

        # 等待搜索框加载完成
        search_input = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "search-bar-input"))
        )

        # 将产品分类名称填入搜索框
        search_input.clear()
        search_input.send_keys(category)

        # 点击搜索按钮
        search_button = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "fy23-icbu-search-bar-inner-button"))
        )
        search_button.click()

        # 等待产品列表加载完成
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "organic-list"))
        )

        # 模拟向下滚动页面，直到加载完所有产品
        last_height = browser.execute_script("return document.body.scrollHeight")
        while True:
            browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)  # 等待加载
            new_height = browser.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # 加载完所有产品后，获取产品列表的长度并打印
        product_list = browser.find_elements(By.CLASS_NAME, "fy23-search-card")
        num_products = len(product_list)
        logging.info(f"共抓取到的产品数量：{num_products}")

        # 循环处理产品
        for product in product_list:
            # 获取产品标题
            product_title = product.find_element(By.CLASS_NAME, "search-card-e-title")
            logging.info(f"当前产品标题: {product_title.text}")

            # 滚动到产品标题所在位置
            scroll_to_element(browser, product_title)
            time.sleep(1)
            # 获取产品链接并打开
            product_link = product.find_element(By.TAG_NAME, "a").get_attribute("href")
            browser.execute_script(f"window.open('{product_link}')")
            # 处理产品详情页操作
            success_count = handle_product_detail(browser, category, success_count)
            # 等待一段时间，可以根据实际情况调整
            time.sleep(1)

        logging.info(f"成功处理的产品数量: {success_count}")
        return success_count

    except Exception as e:
        logging.error(f"处理链接时发生错误: {e}")
        return success_count

def handle_popup(browser):
    try:
        # 等待弹窗出现
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'This product is already in your store, what would you like to do?')]"))
        )

        # 检查是否出现弹窗
        popup_message = browser.find_element(By.XPATH, "//div[contains(text(), 'This product is already in your store, what would you like to do?')]")

        if popup_message:
            logging.info("检测到产品已存在的弹窗")

            # 直接关闭当前产品详情页
            browser.close()
            logging.info("关闭了产品已存在的弹窗")

    except NoSuchElementException:
        logging.info("未检测到产品已存在的弹窗")
    except Exception as e:
        logging.error(f"处理弹窗时发生错误: {e}")

def handle_product_detail(browser, category, success_count):
    try:
        # 获取所有窗口句柄
        original_window = browser.current_window_handle
        handles = browser.window_handles

        # 获取新打开的窗口句柄
        new_window = None
        for window_handle in handles:
            if window_handle != original_window:
                new_window = window_handle
                break

        if new_window:
            # 切换到新打开的产品详情页窗口
            browser.switch_to.window(new_window)

            # 等待产品详情页元素加载完成
            product_title = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.TAG_NAME, "h1")))

            # 滚动到产品标题位置
            scroll_to_element(browser, product_title)

            time.sleep(2)

            # 处理弹窗
            handle_popup(browser)

            # 获取产品标题
            product_title = browser.find_element(By.TAG_NAME, "h1").text
            logging.info(f"产品详情页标题: {product_title}")

            # 处理产品详情页操作，这里可以根据实际需要修改
            success_count = handle_product_actions(browser, category, success_count)

            # 等待一段时间，可以根据实际情况调整
            time.sleep(1)
            # 切换回原始窗口（产品搜索页）
            browser.switch_to.window(original_window)
        return success_count
    except NoSuchWindowException as e:
        logging.error(f"浏览器窗口丢失：{e}")
        return success_count
    except Exception as e:
        logging.error(f"处理产品详情页时发生错误: {e}")
        return success_count

def wait_for_element_to_appear(driver, by, selector, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
    except TimeoutException:
        logging.error(f"元素未能在 {timeout} 秒内出现: {selector}")
        raise

def handle_product_actions(browser, category, success_count):
    try:
        add_btn_con = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="addBtnCon"]')))
        add_btn_con.click()
        logging.info("点击了按钮//*[@id='addBtnCon']")

        try:
            element = WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[@class="inactive" and text()="Draft"]'))
            )
            logging.info("成功加载 Draft 元素")
            actions = ActionChains(browser)
            actions.move_to_element(element).perform()
            element.click()
            logging.info("成功点击 Draft 元素")
            time.sleep(2)
        except Exception as e:
            logging.error(f"等待和点击 Draft 元素时出现错误：{e}")

        time.sleep(3)  # 可以根据实际情况调整等待时间

        # 检查是否出现 "This product is already in your store, what would you like to do?"
        try:
            existing_product_msg = browser.find_element(By.XPATH,
                                                        '//div[contains(text(), "This product is already in your store, what would you like to do?")]')
            logging.info("产品已存在，不再处理当前产品")
            return success_count  # 跳出函数，不再处理当前产品
        except NoSuchElementException:
            pass  # 如果未找到消息元素，继续后续操作

        time.sleep(3)  # 可以根据实际情况调整等待时间

        try:
            # 点击 description_tab_button 按钮
            description_tab_button = browser.find_element(By.XPATH, '//*[@id="description_tab_button"]')
            description_tab_button.click()
            logging.info("点击了 description_tab_button 按钮")
            time.sleep(3)  # 等待页面加载

            # 点击 Variants 按钮
            variants_button = browser.find_element(By.CSS_SELECTOR,
                                                   'button.accordion-tab[data-actab-group="0"][data-actab-id="2"]')
            variants_button.click()
            logging.info("点击了 Variants 按钮")

            # 选择 Import all variants automatically 单选框
            all_variants_radio = browser.find_element(By.ID, 'all_variants')
            all_variants_radio.click()
            logging.info("选择 Import all variants automatically 单选框")

            time.sleep(3)  # 等待页面反应

            # 选择 Select which variants to include 单选框
            price_switch_radio = browser.find_element(By.ID, 'price_switch')
            price_switch_radio.click()
            logging.info("选择 Select which variants to include 单选框")

            time.sleep(3)  # 等待页面反应
        except Exception as e:
            logging.error(f"点击 Variants 按钮时出现错误：{e}")

        # 点击 Images 按钮
        images_button = browser.find_element(By.XPATH,
                                             '//button[@class="accordion-tab accordion-custom-tab" and @data-actab-group="0" and @data-actab-id="3"]')
        images_button.click()
        logging.info("点击了 Images 按钮")
        time.sleep(3)  # 等待页面反应

        add_to_store_button = browser.find_element(By.ID, 'addBtnSec')
        scroll_to_element(browser, add_to_store_button)

        add_to_store_button.click()
        logging.info("成功点击 Add to your Store 按钮")

        logging.info("等待页面加载完成")

        # 等待导入过程完成，确保 importify-app-container 元素出现
        try:
            wait_for_element_to_appear(browser, By.ID, 'importify-app-container')
            logging.info("产品正在导入中...")

            # 等待成功消息出现
            success_message = None
            timeout = 180  # 设定超时时间
            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    success_message = browser.find_element(By.XPATH,
                                                           '//div[@class="textcontainer centeralign home-content "]/p[1]')
                    if success_message.text == "We have successfully created the product page.":
                        logging.info(f"产品导入成功, 共计: {success_count + 1}")
                        success_count += 1
                        break
                    else:
                        logging.warning("产品正在导入中...")
                except Exception as e:
                    logging.warning("未检测到产品成功导入，继续等待...")
                time.sleep(5)  # 每秒检查一次

            if not success_message or success_message.text != "We have successfully created the product page.":
                logging.error("超时：未找到成功创建产品页面的消息")

        except Exception as e:
            logging.error(f"页面加载出错: {e}")
        time.sleep(2)
        # 关闭当前产品详情页标签页
        browser.close()
        time.sleep(1)

        return success_count

    except NoSuchWindowException as e:
        logging.error(f"浏览器窗口丢失：{e}")
        return success_count
    except Exception as e:
        logging.error(f"处理产品详情页操作时发生错误: {e}")
        return success_count

def scroll_to_element(browser, element):
    try:
        WebDriverWait(browser, 10).until(EC.visibility_of(element))
        browser.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        logging.info(f"滚动到元素: {element.text}")
    except Exception as e:
        logging.error(f"滚动到元素时出错: {e}")

def main():
    logging.info("开始主逻辑")
    profile_path = get_valid_profile_path()
    if not profile_path:
        messagebox.showerror("错误", "无效的配置路径。")
        return

    input_product_category(profile_path)

if __name__ == "__main__":
    main()
