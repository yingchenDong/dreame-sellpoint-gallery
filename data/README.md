# 协作数据目录说明

本目录存放网页「协作中心」产生的数据文件，由 GitHub Actions 自动处理，无需手工编辑。

| 路径 | 作用 | 谁写 |
|------|------|------|
| `users.json` | 会员列表：昵称、密码哈希、角色（super / admin / member） | Actions（处理注册与任命） |
| `registrations/` | 注册申请：`reg_<时间戳>_<昵称>.json`，提交后自动成为会员 | 会员（网页生成后上传） |
| `inbox/` | 待审核上传申请：`<id>.json`（申请信息）+ `<id>_<文件名>`（上传文件） | 会员（网页生成后上传） |
| `decisions/` | 审核决定：`<id>.decision.json`，通过则更新网页，拒绝则仅通知 | 管理员（网页审核后上传） |
| `notifications.json` | 通知列表（注册成功 / 审核通过 / 审核拒绝等） | Actions |
| `history.json` | 审核历史记录 | Actions |
| `archive/` | 已处理的申请归档 | Actions |

## 提交位置

- 注册：上传到 `data/registrations/`
- 上传申请：上传到 `data/inbox/`
- 审核决定：上传到 `data/decisions/`

网页「协作中心」会自动生成对应文件并给出上传指引，按指引操作即可。
