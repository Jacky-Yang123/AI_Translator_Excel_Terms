# 修复 SyncManager 初始化错误

## 问题分析

**错误信息**：

```
TypeError: SyncManager.__init__() missing 3 required positional arguments: 'task_id', 'task', and 'logger'
```

**根本原因**：

* 在 `main_window.py` 中，`BackupMainWindow` 类的 `__init__` 方法尝试创建一个 `SyncManager` 实例，但没有传递任何参数

* 然而，在 `backup_sync.py` 中，`SyncManager` 类的构造函数需要三个必需参数：`task_id`, `task`, `logger`

* 相比之下，`MonitorManager` 类的构造函数不需要参数，它在内部管理多个监控器

## 解决方案

### 方案 1：创建 SyncManagerManager 类

创建一个类似于 `MonitorManager` 的类来管理多个 `SyncManager` 实例。

**修改步骤**：

1. **修改 backup\_sync.py**：

   * 创建一个新的 `SyncManagerManager` 类

   * 该类负责管理多个 `SyncManager` 实例

   * 提供添加、移除、获取同步管理器的方法

2. **修改 main\_window\.py**：

   * 将 `self.sync_manager = SyncManager()` 改为 `self.sync_manager_manager = SyncManagerManager()`

   * 更新所有使用 `sync_manager` 的地方，改为使用 `sync_manager_manager`

### 方案 2：修改 SyncManager 类

修改 `SyncManager` 类，使其可以无参数初始化。

**修改步骤**：

1. **修改 backup\_sync.py**：

   * 修改 `SyncManager` 类的构造函数，使其参数变为可选

   * 添加内部方法来管理多个任务

2. **修改 main\_window\.py**：

   * 保持 `self.sync_manager = SyncManager()` 不变

## 推荐方案

**推荐方案 1**：创建 `SyncManagerManager` 类

**理由**：

* 与现有的 `MonitorManager` 设计保持一致

* 更符合面向对象设计原则

* 更易于扩展和维护

* 明确区分单个同步任务和同步任务管理器的职责

## 具体实现步骤

### 1. 在 backup\_sync.py 中添加 SyncManagerManager 类

```python
class SyncManagerManager:
    def __init__(self):
        self.sync_managers: Dict[str, SyncManager] = {}
        self.logger = BackupLogger()

    def add_sync_manager(self, task_id: str, task: Dict) -> SyncManager:
        if task_id in self.sync_managers:
            return self.sync_managers[task_id]
        
        sync_manager = SyncManager(task_id, task, self.logger)
        self.sync_managers[task_id] = sync_manager
        return sync_manager

    def remove_sync_manager(self, task_id: str) -> bool:
        if task_id not in self.sync_managers:
            return False
        
        del self.sync_managers[task_id]
        return True

    def get_sync_manager(self, task_id: str) -> Optional[SyncManager]:
        return self.sync_managers.get(task_id)

    def stop_all(self):
        for sync_manager in self.sync_managers.values():
            # 停止同步管理器的相关操作
            pass
        self.sync_managers.clear()

    def get_all_sync_managers(self) -> Dict[str, SyncManager]:
        return self.sync_managers.copy()
```

### 2. 修改 main\_window\.py

* 将 `self.sync_manager = SyncManager()` 改为 `self.sync_manager_manager = SyncManagerManager()`

* 更新所有使用 `sync_manager` 的地方

### 3. 测试修改

* 运行 `test_modules.py` 验证所有模块都能正常导入

* 运行 `main.py` 启动备份系统

* 测试创建和管理备份任务

## 预期结果

* ✅ 系统能够正常启动，无 `TypeError` 错误

* ✅ 能够创建和管理多个备份任务

* ✅ 能够正常启动和停止同步任务

* ✅ 所有模块测试通过

## 风险评估

* **低风险**：修改只涉及添加新类，不修改现有功能

* **兼容性**：与现有代码完全兼容

* **维护性**：提高了代码的可维护性和扩展性

