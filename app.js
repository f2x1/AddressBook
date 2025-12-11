// 联系人数据管理类
class AddressBook {
    constructor() {
        this.contacts = JSON.parse(localStorage.getItem('contacts')) || [];
        this.currentTab = 'all';
        this.editingId = null;
    }

    init() {
        this.bindEvents();
        this.renderContacts();
    }

    bindEvents() {
        // 切换表单显示
        document.getElementById('toggleFormBtn').addEventListener('click', function() {
            const form = document.getElementById('contactForm');
            form.style.display = form.style.display === 'none' ? 'block' : 'none';
            addressBook.resetForm();
        });

        // 表单提交
        document.getElementById('form').addEventListener('submit', function(e) {
            e.preventDefault();
            addressBook.saveContact();
        });

        // 取消按钮
        document.getElementById('cancelBtn').addEventListener('click', function() {
            document.getElementById('contactForm').style.display = 'none';
            addressBook.resetForm();
        });

        // 添加联系方式
        document.getElementById('addMethodBtn').addEventListener('click', function() {
            addressBook.addContactMethod();
        });

        // 导出Excel
        document.getElementById('exportBtn').addEventListener('click', function() {
            addressBook.exportToExcel();
        });

        // 导入Excel
        document.getElementById('importBtn').addEventListener('change', function(e) {
            addressBook.importFromExcel(e.target.files[0]);
        });

        // 标签页切换
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.addEventListener('click', function(e) {
                addressBook.switchTab(e.target.dataset.tab);
            });
        });

        // 联系方式删除按钮事件委托
        document.getElementById('contactMethods').addEventListener('click', function(e) {
            if (e.target.classList.contains('remove-method')) {
                addressBook.removeContactMethod(e.target.closest('.contact-method'));
            }
        });
    }

    addContactMethod() {
        const contactMethods = document.getElementById('contactMethods');
        const methodDiv = document.createElement('div');
        methodDiv.className = 'contact-method';
        methodDiv.innerHTML = `
            <select class="method-type">
                <option value="phone">电话</option>
                <option value="email">邮箱</option>
                <option value="wechat">微信</option>
                <option value="address">地址</option>
            </select>
            <input type="text" class="method-value" placeholder="请输入联系方式" required>
            <button type="button" class="remove-method">删除</button>
        `;
        contactMethods.appendChild(methodDiv);
    }

    removeContactMethod(methodElement) {
        const contactMethods = document.getElementById('contactMethods');
        if (contactMethods.children.length > 1) {
            contactMethods.removeChild(methodElement);
        } else {
            alert('至少需要一个联系方式');
        }
    }

    getContactMethods() {
        const methods = [];
        document.querySelectorAll('.contact-method').forEach(method => {
            const type = method.querySelector('.method-type').value;
            const value = method.querySelector('.method-value').value;
            if (value) {
                methods.push({ type, value });
            }
        });
        return methods;
    }

    resetForm() {
        document.getElementById('form').reset();
        const contactMethods = document.getElementById('contactMethods');
        // 保留一个联系方式
        while (contactMethods.children.length > 1) {
            contactMethods.removeChild(contactMethods.lastChild);
        }
        this.editingId = null;
    }

    saveContact() {
        const name = document.getElementById('name').value;
        const contactMethods = this.getContactMethods();

        if (!name || contactMethods.length === 0) {
            alert('请填写完整的联系人信息');
            return;
        }

        const contact = {
            name,
            contactMethods,
            isFavorite: false,
            id: this.editingId || Date.now().toString()
        };

        if (this.editingId) {
            // 编辑现有联系人
            const index = this.contacts.findIndex(c => c.id === this.editingId);
            this.contacts[index] = { ...this.contacts[index], ...contact };
        } else {
            // 添加新联系人
            this.contacts.push(contact);
        }

        this.saveToLocalStorage();
        this.renderContacts();
        this.resetForm();
        document.getElementById('contactForm').style.display = 'none';
    }

    toggleFavorite(id) {
        const contact = this.contacts.find(c => c.id === id);
        if (contact) {
            contact.isFavorite = !contact.isFavorite;
            this.saveToLocalStorage();
            this.renderContacts();
        }
    }

    editContact(id) {
        const contact = this.contacts.find(c => c.id === id);
        if (contact) {
            this.editingId = id;
            document.getElementById('name').value = contact.name;
            
            // 清空现有联系方式
            const contactMethods = document.getElementById('contactMethods');
            contactMethods.innerHTML = '';
            
            // 添加联系人的所有联系方式
            contact.contactMethods.forEach(method => {
                const methodDiv = document.createElement('div');
                methodDiv.className = 'contact-method';
                methodDiv.innerHTML = `
                    <select class="method-type">
                        <option value="phone" ${method.type === 'phone' ? 'selected' : ''}>电话</option>
                        <option value="email" ${method.type === 'email' ? 'selected' : ''}>邮箱</option>
                        <option value="wechat" ${method.type === 'wechat' ? 'selected' : ''}>微信</option>
                        <option value="address" ${method.type === 'address' ? 'selected' : ''}>地址</option>
                    </select>
                    <input type="text" class="method-value" value="${method.value}" placeholder="请输入联系方式" required>
                    <button type="button" class="remove-method">删除</button>
                `;
                contactMethods.appendChild(methodDiv);
            });
            
            // 显示表单
            document.getElementById('contactForm').style.display = 'block';
            // 滚动到表单
            document.getElementById('contactForm').scrollIntoView({ behavior: 'smooth' });
        }
    }

    deleteContact(id) {
        if (confirm('确定要删除这个联系人吗？')) {
            this.contacts = this.contacts.filter(c => c.id !== id);
            this.saveToLocalStorage();
            this.renderContacts();
        }
    }

    switchTab(tab) {
        this.currentTab = tab;
        
        // 更新标签页样式
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tab}"]`).classList.add('active');
        
        this.renderContacts();
    }

    renderContacts() {
        const contactList = document.getElementById('contactList');
        let filteredContacts = this.contacts;
        
        if (this.currentTab === 'favorite') {
            filteredContacts = this.contacts.filter(c => c.isFavorite);
        }
        
        if (filteredContacts.length === 0) {
            contactList.innerHTML = '<p style="grid-column: 1 / -1; text-align: center; color: #888; padding: 2rem;">暂无联系人</p>';
            return;
        }
        
        contactList.innerHTML = filteredContacts.map(contact => `
            <div class="contact-card">
                <div class="contact-header">
                    <h3 class="contact-name">${contact.name}</h3>
                    <button class="favorite-button" onclick="addressBook.toggleFavorite('${contact.id}')">
                        ${contact.isFavorite ? '★' : '☆'}
                    </button>
                </div>
                <div class="contact-methods-list">
                    ${contact.contactMethods.map(method => `
                        <div class="contact-method-item">
                            <span class="method-label">${this.getMethodLabel(method.type)}:</span>
                            <span>${method.value}</span>
                        </div>
                    `).join('')}
                </div>
                <div class="card-actions">
                    <button class="edit-btn" onclick="addressBook.editContact('${contact.id}')">编辑</button>
                    <button class="delete-btn" onclick="addressBook.deleteContact('${contact.id}')">删除</button>
                </div>
            </div>
        `).join('');
    }

    getMethodLabel(type) {
        const labels = {
            phone: '电话',
            email: '邮箱',
            wechat: '微信',
            address: '地址'
        };
        return labels[type] || type;
    }

    exportToExcel() {
        if (this.contacts.length === 0) {
            alert('地址簿中没有联系人可导出');
            return;
        }

        // 准备导出数据
        const exportData = this.contacts.map(contact => {
            const row = { 姓名: contact.name };
            
            // 为每种联系方式类型创建列
            contact.contactMethods.forEach(method => {
                const label = this.getMethodLabel(method.type);
                // 如果有多个同类型的联系方式，添加序号
                if (row[label]) {
                    let i = 2;
                    while (row[`${label}${i}`]) i++;
                    row[`${label}${i}`] = method.value;
                } else {
                    row[label] = method.value;
                }
            });
            
            return row;
        });

        // 创建工作簿和工作表
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '联系人');

        // 调整列宽
        const colWidths = Object.keys(exportData[0]).map(key => ({
            wch: Math.max(key.length * 2, 15)
        }));
        ws['!cols'] = colWidths;

        // 导出文件
        XLSX.writeFile(wb, '地址簿导出.xlsx');
    }

    importFromExcel(file) {
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(ws);
                
                this.processImportedData(jsonData);
                alert(`成功导入 ${jsonData.length} 个联系人`);
            } catch (error) {
                alert('导入失败，请检查文件格式是否正确');
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    processImportedData(data) {
        const methodTypes = ['电话', '邮箱', '微信', '地址'];
        
        data.forEach(row => {
            const name = row['姓名'] || row['name'] || '未知姓名';
            const contactMethods = [];
            
            // 处理所有列，提取联系方式
            Object.keys(row).forEach(key => {
                if (key === '姓名' || key === 'name') return;
                
                const value = row[key];
                if (value) {
                    // 匹配联系方式类型
                    let type = 'other';
                    for (const methodType of methodTypes) {
                        if (key.includes(methodType)) {
                            type = methodType === '电话' ? 'phone' :
                                  methodType === '邮箱' ? 'email' :
                                  methodType === '微信' ? 'wechat' : 'address';
                            break;
                        }
                    }
                    
                    contactMethods.push({ type, value });
                }
            });
            
            // 添加到联系人列表
            if (contactMethods.length > 0) {
                this.contacts.push({
                    id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
                    name,
                    contactMethods,
                    isFavorite: false
                });
            }
        });
        
        this.saveToLocalStorage();
        this.renderContacts();
    }

    saveToLocalStorage() {
        localStorage.setItem('contacts', JSON.stringify(this.contacts));
    }
}

// 初始化地址簿
const addressBook = new AddressBook();

// 页面加载完成后执行
window.addEventListener('DOMContentLoaded', () => {
    addressBook.init();
});
