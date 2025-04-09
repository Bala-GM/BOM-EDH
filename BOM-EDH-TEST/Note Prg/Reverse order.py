'''# Define the set
pin_set = {"1-pin", "2-pin", "3-pin", "4-pin", "5-pin", "6-pin", "7-pin", "8-pin", "9-pin", "10-pin"}

# Convert the set to a list
pin_list = list(pin_set)

# Sort the list in descending order
pin_list.sort(reverse=True)

# Print the sorted list
for pin in pin_list:
    print(pin)'''

# Define the set with pin values
pin_set = {
    "1-pin","2-pin","3-pin","4-pin","5-pin","6-pin","7-pin","8-pin","9-pin","10-pin","11-pin","12-pin","13-pin","14-pin","15-pin",
                          "16-pin","17-pin","18-pin","19-pin","20-pin","21-pin","22-pin","23-pin","24-pin","25-pin","26-pin","27-pin","28-pin","29-pin",
                          "30-pin","31-pin","32-pin","33-pin","34-pin","35-pin","36-pin","37-pin","38-pin","39-pin","40-pin","41-pin","42-pin","43-pin",
                          "44-pin","45-pin","46-pin","47-pin","48-pin","49-pin","50-pin","51-pin","52-pin","53-pin","54-pin","55-pin","56-pin","57-pin",
                          "58-pin","59-pin","60-pin","61-pin","62-pin","63-pin","64-pin","65-pin","66-pin","67-pin","68-pin","69-pin","70-pin","71-pin",
                          "72-pin","73-pin","74-pin","75-pin","76-pin","77-pin","78-pin","79-pin","80-pin","81-pin","82-pin","83-pin","84-pin","85-pin",
                          "86-pin","87-pin","88-pin","89-pin","90-pin","91-pin","92-pin","93-pin","94-pin","95-pin","96-pin","97-pin","98-pin","99-pin",
                          "100-pin","101-pin","102-pin","103-pin","104-pin","105-pin","106-pin","107-pin","108-pin","109-pin","110-pin","111-pin","112-pin",
                          "113-pin","114-pin","115-pin","116-pin","117-pin","118-pin","119-pin","120-pin","121-pin","122-pin","123-pin","124-pin","125-pin",
                          "126-pin","127-pin","128-pin","129-pin","130-pin","131-pin","132-pin","133-pin","134-pin","135-pin","136-pin","137-pin","138-pin",
                          "139-pin","140-pin","141-pin","142-pin","143-pin","144-pin","145-pin","146-pin","147-pin","148-pin","149-pin","150-pin","151-pin",
                          "152-pin","153-pin","154-pin","155-pin","156-pin","157-pin","158-pin","159-pin","160-pin","161-pin","162-pin","163-pin","164-pin",
                          "165-pin","166-pin","167-pin","168-pin","169-pin","170-pin","171-pin","172-pin","173-pin","174-pin","175-pin","176-pin","177-pin",
                          "178-pin","179-pin","180-pin","181-pin","182-pin","183-pin","184-pin","185-pin","186-pin","187-pin","188-pin","189-pin","190-pin",
                          "191-pin","192-pin","193-pin","194-pin","195-pin","196-pin","197-pin","198-pin","199-pin","200-pin","201-pin","202-pin","203-pin",
                          "204-pin","205-pin","206-pin","207-pin","208-pin","209-pin","210-pin","211-pin","212-pin","213-pin","214-pin","215-pin","216-pin",
                          "217-pin","218-pin","219-pin","220-pin","221-pin","222-pin","223-pin","224-pin","225-pin","226-pin","227-pin","228-pin","229-pin",
                          "230-pin","231-pin","232-pin","233-pin","234-pin","235-pin","236-pin","237-pin","238-pin","239-pin","240-pin","241-pin","242-pin",
                          "243-pin","244-pin","245-pin","246-pin","247-pin","248-pin","249-pin","250-pin","251-pin","252-pin","253-pin","254-pin","255-pin",
                          "256-pin","257-pin","258-pin","259-pin","260-pin","261-pin","262-pin","263-pin","264-pin","265-pin","266-pin","267-pin","268-pin",
                          "269-pin","270-pin","271-pin","272-pin","273-pin","274-pin","275-pin","276-pin","277-pin","278-pin","279-pin","280-pin","281-pin",
                          "282-pin","283-pin","284-pin","285-pin","286-pin","287-pin","288-pin","289-pin","290-pin","291-pin","292-pin","293-pin","294-pin",
                          "295-pin","296-pin","297-pin","298-pin","299-pin","300-pin"
}

# Convert the set to a list
pin_list = list(pin_set)

# Sort the list based on the numeric part in descending order
pin_list.sort(key=lambda x: int(x.split('-')[0]), reverse=True)

# Print the sorted list
for pin in pin_list:
    print(pin)