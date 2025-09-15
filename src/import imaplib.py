import imaplib
import email
import os
import ssl
from email.header import decode_header
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

def parse_email_date(date_str):
    """парсит дату из письма и возвращает datetime объект"""
    try:
        return parsedate_to_datetime(date_str)
    except:
        try:
            # проба разные форматы дат
            formats = [
                '%a, %d %b %Y %H:%M:%S %z',
                '%a, %d %b %Y %H:%M:%S',
                '%d %b %Y %H:%M:%S %z',
                '%Y-%m-%d %H:%M:%S%z'
            ]
            
            for fmt in formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except:
                    continue
            
            # если не сработало, возвращаем None
            return None
        except:
            return None

def is_date_after_start(email_date, start_date):
    """
    сравнить даты, учитывая часовые пояса
    """
    if email_date is None:
        return False
    
    # приводим обе даты к aware формату с UTC
    if email_date.tzinfo is None:
        email_date = email_date.replace(tzinfo=timezone.utc)
    
    if start_date.tzinfo is None:
        start_date = start_date.replace(tzinfo=timezone.utc)
    
    return email_date >= start_date

def save_mailru_emails_since_date(email_address, password, save_folder):
    """
    сохраняет письма с начиная с 12 сентября 2025 года
    """
    if not save_folder:
        print("папка не указана! Операция отменена.")
        return
    
    # Установка даты - 12 сентября 2025 года (в UTC)
    start_date = datetime(2025, 9, 12, tzinfo=timezone.utc)
    
    abs_path = os.path.abspath(save_folder)
    print(f"письма будут сохранены в: {abs_path}")
    print(f"сохраняем письма начиная с: {start_date.strftime('%d.%m.%Y')}")
    
    if not os.path.exists(abs_path):
        os.makedirs(abs_path)
        print(f"создана папка: {abs_path}")
    
    try:
        context = ssl.create_default_context()
        print("подключаемся к серверу Mail.ru...")
        
        mail = imaplib.IMAP4_SSL("imap.mail.ru", ssl_context=context)
        mail.login(email_address, password)
        mail.select("inbox")
        
        # ищем письма начиная с 12 сентября 2025
        imap_date = start_date.strftime('%d-%b-%Y')
        status, messages = mail.search(None, f'(SINCE "{imap_date}")')
        
        if status != "OK":
            print("ошибка при поиске писем")
            return
        
        email_ids = messages[0].split()
        print(f"найдено писем с {start_date.strftime('%d.%m.%Y')}: {len(email_ids)}")
        
        if not email_ids:
            print("не найдено писем за указанный период")
            return
        
        saved_count = 0
        skipped_count = 0
        
        for i, email_id in enumerate(email_ids):
            try:
                status, msg_data = mail.fetch(email_id, "(RFC822)")
                
                if status != "OK":
                    print(f"Ошибка при получении письма {email_id}")
                    continue
                
                msg = email.message_from_bytes(msg_data[0][1])
                
                # Проверяем дату письма
                email_date_str = msg.get("Date")
                email_date = None
                
                if email_date_str:
                    email_date = parse_email_date(email_date_str)
                    
                    # Если дату не удалось распарсить, пропускаем проверку
                    if email_date is None:
                        print(f"Не удалось распарсить дату письма {email_id}, сохраняем...")
                    elif not is_date_after_start(email_date, start_date):
                        print(f"Пропускаем письмо {email_id} (дата: {email_date_str})")
                        skipped_count += 1
                        continue
                
                # Получаем тему письма
                subject = "Без_темы"
                if msg["Subject"]:
                    subject_header, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject_header, bytes):
                        subject = subject_header.decode(encoding if encoding else "utf-8")
                    else:
                        subject = subject_header
                
                # Получаем отправителя
                from_header = msg.get("From", "Неизвестный отправитель")
                
                # Создаем безопасное имя файла
                safe_subject = "".join(c if c.isalnum() or c in " -_." else "_" for c in subject)
                safe_subject = safe_subject[:50]
                
                # Формируем имя файла с датой и темой
                date_for_filename = email_date.strftime("%Y%m%d") if email_date else "nodate"
                filename = f"{saved_count+1:04d}_{date_for_filename}_{safe_subject}.eml"
                filepath = os.path.join(abs_path, filename)
                
                # Сохраняем письмо
                with open(filepath, "wb") as f:
                    f.write(msg_data[0][1])
                
                print(f"Сохранено: {filename}")
                saved_count += 1
                
            except Exception as e:
                print(f"Ошибка при обработке письма {email_id}: {e}")
                continue
        
        mail.close()
        mail.logout()
        
        print("=" * 50)
        print(f"✓ Обработка завершена!")
        print(f"✓ Найдено писем: {len(email_ids)}")
        print(f"✓ Сохранено писем: {saved_count}")
        print(f"✓ Пропущено (старые даты): {skipped_count}")
        print(f"✓ Письма сохранены в: {abs_path}")
        print("=" * 50)
        
    except Exception as e:
        print(f"Ошибка подключения: {e}")

if __name__ == "__main__":
    print("=" * 60)
    print("СОХРАНЕНИЕ ПИСЕМ С MAIL.RU (С 12.09.2025)")
    print("=" * 60)
    print("ВАЖНО: Используйте пароль для приложений из настроек Mail.ru!")
    print("=" * 60)
    
    email_address = input("Введите ваш email: ").strip()
    password = input("Введите пароль для приложения: ").strip()
    save_folder = input("Введите полный путь к папке для сохранения: ").strip()
    
    if not email_address or not password or not save_folder:
        print("Все поля обязательны для заполнения!")
    else:
        save_mailru_emails_since_date(email_address, password, save_folder)
    
    input("Нажмите Enter для выхода...")