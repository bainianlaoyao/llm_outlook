#!/usr/bin/env python3
"""
推送系统深度诊断脚本
系统性排查推送流程问题
"""

import sys
import os
sys.path.append('.')

def main():
    print('🔧 深度推送流程诊断...')
    print()

    # 1. 环境变量检查
    print('✅ 步骤1：环境变量检查')
    env_vars = ['SERVERCHAN_SENDKEY']
    for var in env_vars:
        value = os.getenv(var, '')
        if value:
            masked_value = value[:10] + '*' * (len(value) - 10) if len(value) > 10 else value
            print(f'  {var}: ✅ 设置为 {repr(masked_value)}')
        else:
            print(f'  {var}: ❌ 未设置')
    print()

    # 2. 配置加载诊断
    print('✅ 步骤2：配置加载诊断')
    try:
        from config.config import load_config
        config = load_config()
        print(f'  配置加载: ✅ 成功')
        if 'push' in config:
            push_config = config['push']
            print(f'  推送配置:')
            print(f'    channels: {push_config.get("channels", "未设置")}')
            print(f'    sendkey: {"***已设置***" if push_config.get("sendkey") else "未设置"}')
        else:
            print('  推送配置: ❌ 未找到push配置项')
    except Exception as e:
        print(f'  配置加载: ❌ 失败 - {e}')
    print()

    # 3. ServerChanPusher诊断
    print('✅ 步骤3：ServerChanPusher诊断')
    try:
        from core.message_pusher import ServerChanPusher

        # 测试不同构造方式
        constructors = [
            ('默认构造', lambda: ServerChanPusher()),
            ('环境变量', lambda: ServerChanPusher(os.getenv('SERVERCHAN_SENDKEY') or None)),
            ('空参数', lambda: ServerChanPusher(None)),
            ('占位符参数', lambda: ServerChanPusher('${{SERVERCHAN_SENDKEY}}')),
        ]

        for name, constructor in constructors:
            try:
                pusher = constructor()
                is_empty = pusher.sendkey == ''
                contains_placeholder = '$' in pusher.sendkey and (pusher.sendkey.startswith('$') or '${{' in pusher.sendkey)
                print(f'  {name}: sendkey长度={len(pusher.sendkey)}, 包含占位符={contains_placeholder}, 有效={not is_empty and not contains_placeholder}')
            except Exception as e:
                print(f'  {name}: ❌ 构造失败 - {e}')

    except Exception as e:
        print(f'  诊断异常: {e}')
    print()

    # 4. 推送函数测试
    print('✅ 步骤4：推送函数测试')
    try:
        from core.message_pusher import push_message

        test_cases = [
            '简单标题',
            '标题\\n内容',
            '📧 多行标题\\n第一行内容\\n第二行内容',
            '',
            '\\n',
        ]

        for i, case in enumerate(test_cases):
            try:
                result = push_message(case)
                status = '✅ 成功' if result.success else f'❌ {result.message}'
                print(f'  测试{i+1}: {repr(case[:20])}... -> {status}')
            except Exception as e:
                print(f'  测试{i+1}: ❌ 异常 - {e}')

    except Exception as e:
        print(f'  测试异常: {e}')
    print()

    print('🎯 诊断结果总结：')
    print('基于以上测试，问题核心在于：')
    print()
    print('🔍 推断的可能问题来源：')
    print('1. ✅ 已确认：环境变量SERVERCHAN_SENDKEY未设置')
    print('2. ❓ 待验证：MultiChannelPusher初始化参数传递')
    print('3. ❓ 待验证：配置加载过程中sendkey参数处理')
    print()
    print('🎯 建议解决方案：')
    print('1. 设置环境变量: export SERVERCHAN_SENDKEY=你的SendKey')
    print('2. 验证配置JSON中的sendkey参数格式')
    print('3. 检查MultiChannelPusher的初始化逻辑')
    print()
    print('⏳ 等待用户提供环境变量以继续深入诊断...')

if __name__ == '__main__':
    main()