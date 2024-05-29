import logging
import threading
import time
import openai
import json
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

CONFIG_FILE = 'config.json'


def load_config():
    """加载配置文件"""
    try:
        with open(CONFIG_FILE) as file:
            return json.load(file)
    except FileNotFoundError:
        return {}


def save_config(config):
    """保存配置文件"""
    with open(CONFIG_FILE, 'w') as file:
        json.dump(config, file)


def get_user_input(prompt, default=None):
    """获取用户输入"""
    if default:
        response = input(f"{prompt} (当前值: {default})，按 Enter 保持不变: ")
        return response if response else default
    else:
        return input(f"{prompt}: ")


def get_config_value(config, key, prompt):
    """获取配置值"""
    if key in config:
        response = get_user_input(f"是否更换 {prompt}?", config[key])
        if response != config[key]:
            config[key] = response
    else:
        config[key] = get_user_input(f"请输入 {prompt}")
    return config[key]


config = load_config()

# 获取和设置配置信息
OPENAI_API_KEY = get_config_value(config, 'openai_api_key', 'OpenAI API 密钥')
OPENAI_API_BASE = get_config_value(config, 'openai_api_base', 'OpenAI API 基础地址')
WORDPRESS_USERNAME = get_config_value(config, 'wordpress_username', 'WordPress 用户名')
WORDPRESS_PASSWORD = get_config_value(config, 'wordpress_password', 'WordPress 密码')
MAX_ITERATIONS = int(get_config_value(config, 'max_iterations', '处理产品的个数'))
save_config(config)

# 设置 OpenAI API 地址
openai.api_base = OPENAI_API_BASE
# 配置 OpenAI API 密钥
openai.api_key = OPENAI_API_KEY

# 产品草稿页面URL
draft_url = "https://www.aggpo.com/wp-admin/edit.php?post_status=draft&post_type=product"


def wait_for_element(driver, by, value, timeout=10):
    """通用等待元素加载方法"""
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        logger.error(f"等待元素加载超时: {value}")
        raise


def wait_until_loaded(driver, timeout=15):
    """等待页面加载完成"""
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, "wpbody-content"))
    )


def login(driver, username, password):
    """登录并等待页面加载完成"""
    driver.get("https://www.aggpo.com/wp-login.php")
    logger.info("导航到登录页面")

    username_input = driver.find_element(By.ID, "user_login")
    password_input = driver.find_element(By.ID, "user_pass")
    submit_button = driver.find_element(By.ID, "wp-submit")

    username_input.send_keys(username)
    password_input.send_keys(password)
    submit_button.click()
    logger.info("成功填入账户信息，正在登录...")

    wait_until_loaded(driver)  # 等待页面加载完成


def open_draft_edit_page(driver, draft_link):
    """打开草稿编辑页面"""
    driver.get(draft_link)
    logger.info("打开草稿编辑页面")
    wait_until_loaded(driver)  # 等待页面加载完成


def get_product_title(driver):
    """获取产品标题"""
    try:
        wait_for_element(driver, By.XPATH, '//*[@id="title"]')
        product_title = driver.find_element(By.XPATH, '//*[@id="title"]').get_attribute("value")
        return product_title
    except Exception as e:
        logger.error(f"获取产品标题时发生错误：{e}")
        return None


def scroll_to_element_by_xpath(driver, xpath):
    """通过XPath滚动到指定元素"""
    try:
        # 等待元素加载完成
        wait_for_element(driver, By.XPATH, xpath)
        # 使用JavaScript将页面滚动到指定元素
        scroll_script = (f"document.evaluate('{xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, "
                         f"null).singleNodeValue.scrollIntoView({{ behavior: 'smooth', block: 'center' }});")
        driver.execute_script(scroll_script)
        logger.info(f"已滚动到元素 XPath: {xpath} 并保持在页面中间")
    except Exception as e:
        logger.error(f"滚动到指定元素时发生错误：{e}")


def generate_content(product_title, prompt, max_retries=10, max_tokens=50):
    """使用 OpenAI 生成关键词或描述"""
    result = None  # 默认备选方案移除
    success = False  # 默认标志为失败
    attempts = 0  # 初始化尝试次数

    def run_openai():
        nonlocal result, success, attempts
        while attempts < max_retries and not success:
            try:
                response = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": "You are a seo helpful assistant."},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=max_tokens,
                    temperature=0.7,
                    timeout=60
                )
                raw_result = response.choices[0].message['content'].strip().replace('"', '')  # 移除双引号
                # 确保返回值不超过指定的字符数并截取到句子结尾
                if len(raw_result) > max_tokens:
                    truncated_result = raw_result[:max_tokens]
                    last_period_index = truncated_result.rfind('.')
                    if last_period_index != -1:
                        result = truncated_result[:last_period_index + 1]
                    else:
                        result = truncated_result
                else:
                    result = raw_result
                success = True  # 设置成功标志为 True
            except Exception as e:
                attempts += 1  # 增加尝试次数
                logger.warning(f"OpenAI 请求失败（尝试次数 {attempts}/{max_retries}）：{e}")

    t = threading.Thread(target=run_openai)
    t.start()
    t.join(timeout=60 * max_retries)  # 等待线程执行，最大超时时间根据重试次数

    return result, success  # 返回结果和成功标志


def fill_keywords(driver, product_title):
    """填充关键词"""
    generated_keyword, success = generate_content(product_title,
                                                  f"Provide a concise SEO keyphrase for the product {product_title} that is beneficial for SEO. within 50 characters.",
                                                  max_tokens=50)

    if not success:
        logger.warning(f"未能成功生成关键词：{product_title}")
        return False  # 返回失败标志

    try:
        wait_for_element(driver, By.ID, "focus-keyword-input-metabox")
        keyword_input = driver.find_element(By.ID, "focus-keyword-input-metabox")
        keyword_input.clear()
        keyword_input.send_keys(generated_keyword)
        logger.info(f"成功获取 OpenAI 返回值：生成的关键词为 {generated_keyword}")
        return True  # 返回成功标志

    except Exception as e:
        logger.error(f"填充关键词时发生错误：{e}")
        return False  # 返回失败标志


def fill_description(driver, product_title):
    """填充描述"""
    generated_description, success = generate_content(product_title,
                                                      f"Generate a compelling description for the product {product_title} within 150 characters.",
                                                      max_retries=5, max_tokens=150)

    if not success:
        logger.warning(f"未能成功生成描述：{product_title}")
        return False  # 返回失败标志

    try:
        wait_for_element(driver, By.ID, "yoast-google-preview-description-metabox")
        description_input = driver.find_element(By.ID, "yoast-google-preview-description-metabox")
        description_input.clear()
        description_input.send_keys(generated_description)
        logger.info(f"成功获取 OpenAI 返回值：生成的描述为 {generated_description}")
        return True  # 返回成功标志

    except Exception as e:
        logger.error(f"填充描述时发生错误：{e}")
        return False  # 返回失败标志


def scroll_to_woocommerce_product_data(driver):
    """滚动页面至 woocommerce-product-data 元素"""
    try:
        wait_for_element(driver, By.ID, "woocommerce-product-data")
        product_data_element = driver.find_element(By.ID, "woocommerce-product-data")
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });",
                              product_data_element)
        logger.info("已将 woocommerce-product-data 元素滚动到页面中间")
    except Exception as e:
        logger.error(f"滚动至 woocommerce-product-data 元素时发生错误：{e}")


def click_product_type(driver):
    """点击产品类型选项卡"""
    try:
        wait_for_element(driver, By.ID, "product-type")
        product_type_tab = driver.find_element(By.ID, "product-type")
        product_type_tab.click()
        logger.info("成功点击产品类型选项卡")
    except Exception as e:
        logger.error(f"点击产品类型选项卡时发生错误：{e}")


def select_variable_product(driver):
    """选择 Variable product"""
    try:
        wait_for_element(driver, By.ID, "product-type")
        product_type_select = Select(driver.find_element(By.ID, "product-type"))
        product_type_select.select_by_value("variable")
        logger.info("成功选择产品类型为 'Variable product'")
    except Exception as e:
        logger.error(f"选择产品类型时发生错误：{e}")


def scroll_to_variations_tab(driver):
    """滚动页面至 Variations 标签"""
    try:
        wait_for_element(driver, By.CSS_SELECTOR, "li.variations_options.variations_tab.show_if_variable")
        variations_tab = driver.find_element(By.CSS_SELECTOR, "li.variations_options.variations_tab.show_if_variable")
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", variations_tab)
        logger.info("已将 Variations 标签滚动到页面中间")
    except Exception as e:
        logger.error(f"滚动到 Variations 标签时发生错误：{e}")


def click_variations_tab(driver):
    """点击 Variations 标签"""
    try:
        wait_for_element(driver, By.CSS_SELECTOR, "li.variations_options.variations_tab.show_if_variable a")
        variations_tab = driver.find_element(By.CSS_SELECTOR, "li.variations_options.variations_tab.show_if_variable a")
        variations_tab.click()
        logger.info("成功点击 Variations 标签")
    except Exception as e:
        logger.error(f"点击 Variations 标签时发生错误：{e}")


def check_edit_variation_element(driver):
    """检查 Edit Variation 元素是否存在"""
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.edit_variation.edit"))
        )
        edit_variation_element = driver.find_element(By.CSS_SELECTOR, "a.edit_variation.edit")
        logger.info("找到 Edit Variation 元素")
        return edit_variation_element
    except TimeoutException:
        logger.info("未找到 Edit Variation 元素")
        return None


def scroll_to_edit_variation(driver):
    """滚动页面至 Edit Variation 元素"""
    try:
        wait_for_element(driver, By.CSS_SELECTOR, "a.edit_variation.edit")
        edit_variation_element = driver.find_element(By.CSS_SELECTOR, "a.edit_variation.edit")
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });",
                              edit_variation_element)
        logger.info("已将 Edit Variation 元素滚动到页面中间")
    except Exception as e:
        logger.error(f"滚动至 Edit Variation 元素时发生错误：{e}")


def click_edit_variation(driver):
    """点击 Edit Variation 链接"""
    try:
        wait_for_element(driver, By.CSS_SELECTOR, "a.edit_variation.edit")
        edit_variation_element = driver.find_element(By.CSS_SELECTOR, "a.edit_variation.edit")
        edit_variation_element.click()
        logger.info("成功点击 Edit Variation 链接")
    except Exception as e:
        logger.error(f"点击 Edit Variation 链接时发生错误：{e}")


def scroll_to_variable_regular_price(driver):
    """滚动页面至 variable_regular_price_0 元素"""
    try:
        wait_for_element(driver, By.ID, "variable_regular_price_0")
        variable_price_element = driver.find_element(By.ID, "variable_regular_price_0")
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });",
                              variable_price_element)
        logger.info("已将 variable_regular_price_0 元素滚动到页面中间")
    except Exception as e:
        logger.error(f"滚动至 variable_regular_price_0 元素时发生错误：{e}")


def copy_variable_regular_price(driver):
    """复制 variable_regular_price_0 的价格值"""
    try:
        wait_for_element(driver, By.ID, "variable_regular_price_0")
        variable_regular_price = driver.find_element(By.ID, "variable_regular_price_0")
        price_value = variable_regular_price.get_attribute("value")
        if price_value is None:
            raise ValueError("未能获取 variable_regular_price_0 的价格值")
        logger.info(f"成功复制 variable_regular_price_0 的价格值为：{price_value}")
        return price_value
    except Exception as e:
        logger.error(f"复制 variable_regular_price_0 的价格值时发生错误：{e}")
        return None


def paste_price_to_regular(driver, price):
    """将价格值填入 _regular_price"""
    try:
        wait_for_element(driver, By.ID, "_regular_price")
        regular_price_input = driver.find_element(By.ID, "_regular_price")
        regular_price_input.clear()
        regular_price_input.send_keys(price)
        logger.info(f"成功将价格值 {price} 填入 _regular_price")
    except Exception as e:
        logger.error(f"填入价格值时发生错误：{e}")


def paste_price_to_max_range(driver, price_value):
    """将价格值的两倍填入 _max_price_for_range"""
    try:
        # 转换价格值为浮点数或整数
        price_float = float(price_value)  # 或者使用 int(price_value) 如果价格值是整数形式的
        # 计算价格值的两倍
        double_price = price_float * 2
        # 在 _max_price_for_range 输入框中填入两倍价格值
        wait_for_element(driver, By.ID, "_max_price_for_range")
        max_price_input = driver.find_element(By.ID, "_max_price_for_range")
        max_price_input.clear()
        max_price_input.send_keys(str(double_price))  # 将结果转换为字符串输入
        logger.info(f"成功将价格值的两倍填入 _max_price_for_range：{double_price}")
    except Exception as e:
        logger.error(f"填入价格值的两倍时发生错误：{e}")


def scroll_to_top(driver):
    """滚动页面至顶部"""
    try:
        driver.execute_script("window.scrollTo(0, 0)")
        logger.info("成功滚动至页面顶部")
    except Exception as e:
        logger.error(f"滚动至页面顶部时发生错误：{e}")


def wait_for_element(driver, by, value, timeout=30):
    """等待元素出现"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except Exception as e:
        logger.error(f"等待元素时发生错误：{e}")
        raise


def publish_product(driver):
    """发布产品"""
    try:
        wait_for_element(driver, By.XPATH, '//*[@id="publish"]')
        publish_button = driver.find_element(By.XPATH, '//*[@id="publish"]')
        publish_button.click()

        # 等待成功消息出现
        success_message = None
        timeout = 60  # 设定超时时间
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                success_message = driver.find_element(By.XPATH, '//*[@id="message" and contains(@class, "notice-success")]')
                if success_message:
                    logger.info("产品发布成功")
                    break
                else:
                    logger.warning("产品正在发布中...")
            except Exception as e:
                logger.warning("未检测到产品成功发布消息，继续等待...")
            time.sleep(5)  # 每5秒检查一次

        if not success_message:
            logger.error("超时：未找到成功发布产品的消息")
            return False

        return True

    except Exception as e:
        logger.error(f"发布产品时发生错误：{e}")
        return False


def select_simple_product(driver):
    """选择 Simple product"""
    try:
        wait_for_element(driver, By.ID, "product-type")
        product_type_select = Select(driver.find_element(By.ID, "product-type"))
        product_type_select.select_by_value("simple")
        logger.info("成功选择产品类型为 'Simple product'")
    except Exception as e:
        logger.error(f"选择产品类型时发生错误：{e}")


def scroll_to_regular_price(driver):
    """滚动页面至 Regular Price"""
    try:
        wait_for_element(driver, By.ID, "_regular_price")
        regular_price_element = driver.find_element(By.ID, "_regular_price")
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });",
                              regular_price_element)
        logger.info("已将 Regular Price 元素滚动到页面中间")
    except Exception as e:
        logger.error(f"滚动至 Regular Price 元素时发生错误：{e}")


def copy_regular_price(driver):
    """复制 Regular Price 的价格值"""
    try:
        wait_for_element(driver, By.ID, "_regular_price")
        regular_price_input = driver.find_element(By.ID, "_regular_price")
        price_value = regular_price_input.get_attribute("value")
        if price_value is None:
            raise ValueError("未能获取 _regular_price 的价格值")
        logger.info(f"成功复制 Regular Price 的价格值为：{price_value}")
        return price_value
    except Exception as e:
        logger.error(f"复制 Regular Price 的价格值时发生错误：{e}")
        return None


def remove_blocking_overlay(driver):
    """移除任何阻塞的 overlay 元素"""
    try:
        driver.execute_script("document.querySelectorAll('.blockUI.blockOverlay').forEach(el => el.remove());")
        logger.info("成功移除阻塞的 overlay 元素")
    except Exception as e:
        logger.error(f"移除阻塞的 overlay 元素时发生错误：{e}")


def process_drafts(driver):
    drafts_processed = 0
    max_iterations = MAX_ITERATIONS  # 设置循环次数

    while drafts_processed < max_iterations:
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "row-title")))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            drafts = soup.find_all("a", class_="row-title")

            if len(drafts) == 0:
                logger.info("没有更多产品草稿，退出循环")
                break

            for draft in drafts:
                draft_link = draft["href"]
                open_draft_edit_page(driver, draft_link)  # 打开草稿编辑页面
                product_title = get_product_title(driver)  # 获取产品标题
                time.sleep(2)  # 停留1秒
                scroll_to_element_by_xpath(driver, '//*[@id="focus-keyword-input-metabox"]')  # 滚动到关键词输入框
                if not fill_keywords(driver, product_title):  # 填充关键词，如果失败则跳过草稿
                    logger.warning(f"跳过草稿 {product_title}，关键词生成失败")
                    continue
                time.sleep(2)  # 停留1秒
                scroll_to_element_by_xpath(driver, '//*[@id="yoast-google-preview-description-metabox"]')  # 滚动到描述输入框
                if not fill_description(driver, product_title):  # 填充描述，如果失败则跳过草稿
                    logger.warning(f"跳过草稿 {product_title}，描述生成失败")
                    continue
                time.sleep(2)  # 停留1秒
                scroll_to_woocommerce_product_data(driver)  # 滚动至 woocommerce-product-data 元素
                time.sleep(1)  # 停留1秒
                click_product_type(driver)  # 点击产品类型选项卡
                time.sleep(1)  # 停留1秒
                select_variable_product(driver)  # 选择Variable product
                time.sleep(1)  # 停留1秒
                scroll_to_variations_tab(driver)  # 在处理草稿的过程中调用滚动到 Variations 选项卡的函数
                time.sleep(1)  # 停留1秒
                click_variations_tab(driver)  # 在处理草稿的过程中调用点击 Variations 选项卡的函数
                time.sleep(2)  # 停留1秒

                # 检查 Edit Variation 元素是否存在
                edit_variation_element = check_edit_variation_element(driver)
                if edit_variation_element:
                    scroll_to_edit_variation(driver)  # 滚动到 Edit Variation 元素
                    time.sleep(1)  # 停留1秒
                    click_edit_variation(driver)  # 点击 Edit Variation 链接
                    time.sleep(2)  # 停留1秒
                    scroll_to_variable_regular_price(driver)  # 滚动至 variable_regular_price_0 元素
                    time.sleep(1)  # 停留1秒
                    price_value = copy_variable_regular_price(driver)  # 复制variable_regular_price_0的价格值
                    time.sleep(1)  # 停留1秒
                    scroll_to_woocommerce_product_data(driver)  # 滚动至 woocommerce-product-data 元素
                    time.sleep(1)  # 停留1秒
                    click_product_type(driver)  # 点击产品类型选项卡
                    time.sleep(1)  # 停留1秒
                    select_simple_product(driver)  # 选择Simple product
                    time.sleep(2)  # 停留1秒
                    scroll_to_regular_price(driver)  # 滚动至 _regular_price 元素
                    time.sleep(1)  # 停留1秒
                    paste_price_to_regular(driver, price_value)  # 将价格值填入_regular_price
                    time.sleep(1)  # 停留1秒
                    paste_price_to_max_range(driver, price_value)  # 将价格值的两倍填入_max_price_for_range
                    time.sleep(1)  # 停留1秒
                else:
                    logger.info("未找到 Edit Variation 元素，执行备用操作")
                    scroll_to_woocommerce_product_data(driver)  # 滚动至 woocommerce-product-data 元素
                    time.sleep(1)  # 停留1秒
                    click_product_type(driver)  # 点击产品类型选项卡
                    time.sleep(1)  # 停留1秒
                    select_simple_product(driver)  # 选择Simple product
                    time.sleep(2)  # 停留2秒
                    scroll_to_regular_price(driver)  # 滚动至 _regular_price 元素
                    time.sleep(1)  # 停留1秒

                    # 检查 _regular_price 元素是否有值
                    try:
                        regular_price_input = driver.find_element(By.ID, "_regular_price")
                        regular_price_value = regular_price_input.get_attribute("value")
                        if regular_price_value:
                            paste_price_to_max_range(driver, regular_price_value)  # 将价格值的两倍填入_max_price_for_range
                        else:
                            logger.info("未找到 Regular Price 值，跳过该草稿")
                            continue  # 跳过该草稿
                    except NoSuchElementException as e:
                        logger.error(f"未找到 Regular Price 输入框：{e}")
                        continue  # 跳过该草稿

                scroll_to_top(driver)  # 滚动页面至顶部
                time.sleep(1)  # 停留1秒
                publish_product(driver)  # 发布产品
                time.sleep(1)  # 停留3秒

                drafts_processed += 1
                logger.info(f"成功处理第 {drafts_processed} 个产品草稿")

            driver.get(draft_url)
        except Exception as e:
            logger.error(f"发生错误：{e}")
            continue  # 继续处理下一个草稿

    return drafts_processed


def main():
    driver = None

    try:
        driver = webdriver.Firefox()

        login(driver, WORDPRESS_USERNAME, WORDPRESS_PASSWORD)  # 调用登录函数并传递用户名和密码
        logger.info("登录成功，开始处理产品草稿")

        driver.get(draft_url)  # 导航到产品草稿页面
        drafts_processed = process_drafts(driver)  # 处理产品草稿
        logger.info(f"成功处理 {drafts_processed} 个产品草稿")

    except Exception as e:
        logger.error(f"发生错误：{e}")
    finally:
        if driver is not None:
            driver.quit()
            logger.info("关闭浏览器")


if __name__ == "__main__":
    main()
