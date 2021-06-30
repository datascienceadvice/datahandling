d = openxlsx::read.xlsx("examples/missed.cases.example.xlsx", sheet = "visits")

# при подсчете количественных показателей пропущенные наблюдения, как правило,
# не вызывают особых сложностей
# основная проблема заключается в соотнесении пропущенного наблюдения и
# нахождения долей

# предположим, что задача исследователя заключается в определении частоты
# жалоб пациентов в различные дни исследования

# День 1
no.complaints <- sum(d$complaints[d$visit == "День 1"] == "нет")
minor.complaints <- sum(d$complaints[d$visit == "День 1"] == "незначительные")
n.patients    <- length(d$complaints[d$visit == "День 1"])
sprintf("Жалоб нет: %.2f%%", no.complaints/n.patients * 100)
sprintf("Незначительные жалобы: %.2f%%", minor.complaints/n.patients * 100)

# День 14
no.complaints <- sum(d$complaints[d$visit == "День 14"] == "нет")
minor.complaints <- sum(d$complaints[d$visit == "День 14"] == "незначительные")
n.patients    <- length(d$complaints[d$visit == "День 14"])
sprintf("Жалоб нет: %.2f%%", no.complaints/n.patients * 100)
sprintf("Незначительные жалобы: %.2f%%", minor.complaints/n.patients * 100)

# День 21
no.complaints <- sum(d$complaints[d$visit == "День 21"] == "нет")
minor.complaints <- sum(d$complaints[d$visit == "День 21"] == "незначительные")
n.patients    <- length(d$complaints[d$visit == "День 21"])
sprintf("Жалоб нет: %.2f%%", no.complaints/n.patients * 100)
sprintf("Незначительные жалобы: %.2f%%", minor.complaints/n.patients * 100)

# без учета того, что один из пациентов пропустил визит на 14 день можно
# не совсем корректно заключить, что 100% пациентов имели незначительные жалобы
# решением является указание не только доли, но и количества пациентов
# а также учет пациентов, пропустивших визит

# День 14
no.complaints <- sum(d$complaints[d$visit == "День 14"] == "нет")
minor.complaints <- sum(d$complaints[d$visit == "День 14"] == "незначительные")
n.patients    <- length(d$complaints[d$visit == "День 14"])
sprintf("Жалоб нет: %.2f%% (%i/%i)", no.complaints/n.patients * 100, no.complaints, n.patients)
sprintf("Незначительные жалобы: %.2f%% (%i/%i)", minor.complaints/n.patients * 100, minor.complaints, n.patients)

# количество пропущенных визитов
# предполагаем, что все пациенты в исследовании посетили первый визит
ids <- d$id[d$visit == "День 1"]
missed.visit <- sum(!ids %in% d$id[d$visit == "День 14"])
n.patients <- length(ids)
sprintf("Визит пропущен: %.2f%% (%i/%i)", missed.visit/n.patients * 100, missed.visit, n.patients)

# иногда, однако, каждому из пациентов может принадлежать более одного наблюдения
# как в случае, например, с незапланированными визитами
# поэтому стоит отдельно находить количество незапланированных визитов,
# количство пациентов, совершивших незапланированные визиты и учитывать это
# при анализе частоты жалоб

# количество незапланированных визитов
du = openxlsx::read.xlsx("examples/missed.cases.example.xlsx", sheet = "visits.unplanned")

sprintf("Количество незапланированных визитов: %i", nrow(du))
sprintf("Количство пациентов, совершивших незапланированные визиты: %i", length(unique(du$id)))

aggregate(complaints ~ id, du, length) # незапланированные визиты по пациентам

# количестов пациентов с жалобами
compl.patients <- length(unique(du$id))
n.patients <- length(ids)
sprintf("Пациенты с жалобами: %.2f%% (%i/%i)", compl.patients/n.patients * 100, compl.patients, n.patients)

# жалобы по пациентам
table(du$id, du$complaints)

headache   <- length(unique(du$id[du$complaints == "головная боль"]))
vomiting   <- length(unique(du$id[du$complaints == "рвота"]))
fever      <- length(unique(du$id[du$complaints == "жар"]))
n.patients <- length(unique(du$id))

sprintf("Головная боль: %.2f%% (%i/%i)", headache/n.patients * 100, headache, n.patients)
sprintf("Рвота: %.2f%% (%i/%i)", vomiting/n.patients * 100, vomiting, n.patients)
sprintf("Жар: %.2f%% (%i/%i)", fever/n.patients * 100, fever, n.patients)


