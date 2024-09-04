from win32com.client import Dispatch

L = Dispatch('LEAP.LEAPApplication')
p = L.Branch("\Key Assumptions\AB COM")

dummyInput = "aa"
dummyOutput = "ad"
dummyFile = "ab com.xlsx"  # Case sensitive


# 6 sub categories

for b in p.Children:
    print(b.Name)
    if b.BranchType == 9:
        print("category", "  ", b.Name)
        for n in b.Children:
            if n.BranchType == 10:
                m1 = n.Variables("Key Assumptions")
                print(n.name, "  ", m1.Expression)
                if dummyFile in m1.Expression:
                    exp = m1.Expression.replace(dummyInput, dummyOutput)
                    m1.Expression = exp
                print(n.name, "  ", m1.Expression)
            if n.BranchType == 9:
                print("subcategory", "  ", n.Name)
                for n1 in n.children:
                    if n1.BranchType == 10:
                        m2 = n1.Variables("Key Assumptions")
                        print(n1.name, "  ", m2.Expression)
                        if dummyFile in m2.Expression:
                            exp = m2.Expression.replace(dummyInput, dummyOutput)
                            m2.Expression = exp
                        print(n1.name, "  ", m2.Expression)
                    if n1.BranchType == 9:
                        print("subcategory 2", "  ", n1.Name)
                        for n2 in n1.children:
                            if n2.BranchType == 10:
                                m3 = n2.Variables("Key Assumptions")
                                print(n2.name, "  ", m3.Expression)
                                if dummyFile in m3.Expression:
                                    exp = m3.Expression.replace(dummyInput, dummyOutput)
                                    m3.Expression = exp
                                print(n2.name, "  ", m3.Expression)
                            if n2.BranchType == 9:
                                print("subcategory 3", "  ", n2.Name)
                                for n3 in n2.children:
                                    if n3.BranchType == 10:
                                        m4 = n3.Variables("Key Assumptions")
                                        print(n3.name, "  ", m4.Expression)
                                        if dummyFile in m4.Expression:
                                            exp = m4.Expression.replace(dummyInput, dummyOutput)
                                            m4.Expression = exp
                                        print(n3.name, "  ", m4.Expression)
                                    if n3.BranchType == 9:
                                        print("subcategory 4", "  ", n3.Name)
                                        for n4 in n3.children:
                                            if n4.BranchType == 10:
                                                m5 = n4.Variables("Key Assumptions")
                                                print(n4.name, "  ", m5.Expression)
                                                if dummyFile in m5.Expression:
                                                    exp = m5.Expression.replace(dummyInput, dummyOutput)
                                                    m5.Expression = exp
                                                print(n4.name, "  ", m5.Expression)
                                            if n4.BranchType == 9:
                                                print("subcategory 5", "  ", n4.Name)
                                                for n5 in n4.children:
                                                    if n5.BranchType == 10:
                                                        m6 = n5.Variables("Key Assumptions")
                                                        print(n5.name, "  ", m6.Expression)
                                                        if dummyFile in m6.Expression:
                                                            exp = m6.Expression.replace(dummyInput, dummyOutput)
                                                            m6.Expression = exp
                                                        print(n5.name, "  ", m6.Expression)
                                                    if n5.BranchType == 9:
                                                        print("subcategory 6", "  ", n5.Name)
                                                        for n6 in n5.children:
                                                            if n6.BranchType == 10:
                                                                m7 = n6.Variables("Key Assumptions")
                                                                print(n6.name, "  ", m7.Expression)
                                                                if dummyFile in m7.Expression:
                                                                    exp = m7.Expression.replace(dummyInput, dummyOutput)
                                                                    m7.Expression = exp
                                                                print(n6.name, "  ", m7.Expression)
                                                            if n6.BranchType == 9:
                                                                print("subcategory 7", "  ", n6.Name)

    elif b.BranchType == 10:
        v = b.Variables("Key Assumptions")
        print(b.Name, "  ", v.Expression)
        if dummyFile in v.Expression:
            exp = v.Expression.replace(dummyInput, dummyOutput)
            v.Expression = exp
        print(b.Name, "  ", v.Expression)
    else:
        print("Oh no a ghost")
